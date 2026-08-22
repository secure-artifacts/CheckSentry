<#
.SYNOPSIS
    从 Windows 卸载注册表读取已安装软件。显式读取 32/64 位注册表视图，
    不使用 Win32_Product，也不依赖当前 PowerShell 进程位数。
#>

function Expand-RegistryPathValue {
    param([object]$Value)
    if ($null -eq $Value) { return '' }
    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return '' }
    return [Environment]::ExpandEnvironmentVariables($text)
}

function Resolve-DisplayIconPath {
    param([object]$DisplayIcon)
    $value = Expand-RegistryPathValue -Value $DisplayIcon
    if ([string]::IsNullOrWhiteSpace($value)) { return '' }

    $candidate = $value
    if ($value -match '^"([^"]+)"') {
        $candidate = $matches[1]
    } elseif ($value -match '^(.+?\.(?:exe|dll|ico|png|jpg|jpeg|gif))(?:,\s*-?\d+)?(?:\s+.*)?$') {
        $candidate = $matches[1]
    }

    $candidate = $candidate.Trim().Trim('"')
    if (Test-Path -LiteralPath $candidate -PathType Leaf) {
        $item = Get-Item -LiteralPath $candidate -ErrorAction SilentlyContinue
        if ($null -ne $item) { return $item.FullName }
    }
    return ''
}

function Resolve-InstallLocationIconPath {
    param(
        [string]$InstallLocation,
        [string]$DisplayName
    )
    if ([string]::IsNullOrWhiteSpace($InstallLocation)) { return '' }
    $base = (Expand-RegistryPathValue -Value $InstallLocation).Trim().Trim('"')
    if (-not (Test-Path -LiteralPath $base -PathType Container)) { return '' }

    $candidates = @()
    $name = [string]$DisplayName
    if ($name -like '*LibreOffice*') {
        $candidates += (Join-Path $base 'program\soffice.exe')
        $candidates += (Join-Path $base 'program\swriter.exe')
    }
    if ($name -like '*VMware Tools*') {
        $candidates += (Join-Path $base 'vmtoolsd.exe')
    }

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            $item = Get-Item -LiteralPath $candidate -ErrorAction SilentlyContinue
            if ($null -ne $item) { return $item.FullName }
        }
    }

    try {
        $fallback = Get-ChildItem -LiteralPath $base -Filter '*.exe' -File -ErrorAction Stop |
            Where-Object { $_.Name -notmatch '^(unins|uninstall|setup|update|crash|helper)' } |
            Select-Object -First 1
        if ($null -ne $fallback) { return $fallback.FullName }
    } catch {
        Write-Verbose "枚举安装目录中的图标候选程序失败：$base。$($_.Exception.Message)"
    }
    return ''
}

function Get-RegistryUninstallEntries {
    param(
        [Microsoft.Win32.RegistryHive]$Hive,
        [Microsoft.Win32.RegistryView]$View
    )

    $baseKey = $null
    $uninstallKey = $null
    try {
        $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey($Hive, $View)
        $uninstallKey = $baseKey.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall', $false)
        if ($null -eq $uninstallKey) { return @() }

        $items = @()
        foreach ($subKeyName in $uninstallKey.GetSubKeyNames()) {
            $subKey = $null
            try {
                $subKey = $uninstallKey.OpenSubKey($subKeyName, $false)
                if ($null -eq $subKey) { continue }
                $options = [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames
                $items += [PSCustomObject]@{
                    DisplayName     = $subKey.GetValue('DisplayName', $null, $options)
                    DisplayVersion  = $subKey.GetValue('DisplayVersion', $null, $options)
                    Publisher       = $subKey.GetValue('Publisher', $null, $options)
                    InstallDate     = $subKey.GetValue('InstallDate', $null, $options)
                    InstallLocation = $subKey.GetValue('InstallLocation', $null, $options)
                    UninstallString = $subKey.GetValue('UninstallString', $null, $options)
                    DisplayIcon     = $subKey.GetValue('DisplayIcon', $null, $options)
                    SystemComponent = $subKey.GetValue('SystemComponent', $null, $options)
                    ParentKeyName   = $subKey.GetValue('ParentKeyName', $null, $options)
                    ReleaseType     = $subKey.GetValue('ReleaseType', $null, $options)
                    RegistryHive    = [string]$Hive
                    RegistryView    = [string]$View
                    RegistryKeyName = $subKeyName
                }
            } catch {
                continue
            } finally {
                if ($null -ne $subKey) { $subKey.Dispose() }
            }
        }
        return $items
    } finally {
        if ($null -ne $uninstallKey) { $uninstallKey.Dispose() }
        if ($null -ne $baseKey) { $baseKey.Dispose() }
    }
}

function Get-InstalledSoftwareList {
    [CmdletBinding()]
    param([switch]$IncludeSystemComponents)

    $views = @([Microsoft.Win32.RegistryView]::Default)
    if ([Environment]::Is64BitOperatingSystem) {
        $views = @(
            [Microsoft.Win32.RegistryView]::Registry64,
            [Microsoft.Win32.RegistryView]::Registry32
        )
    }

    $raw = @()
    foreach ($hive in @([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryHive]::CurrentUser)) {
        foreach ($view in $views) {
            $raw += @(Get-RegistryUninstallEntries -Hive $hive -View $view)
        }
    }

    $raw = @($raw | Where-Object {
        $null -ne $_.DisplayName -and -not [string]::IsNullOrWhiteSpace([string]$_.DisplayName)
    })

    if (-not $IncludeSystemComponents) {
        $raw = @($raw | Where-Object {
            $_.SystemComponent -ne 1 -and
            [string]::IsNullOrWhiteSpace([string]$_.ParentKeyName) -and
            ([string]$_.ReleaseType -notin @('Update', 'Hotfix', 'ServicePack', 'Security Update'))
        })
    }

    $result = foreach ($item in $raw) {
        $displayName = ([string]$item.DisplayName).Trim()
        $installLocation = (Expand-RegistryPathValue -Value $item.InstallLocation).Trim().Trim('"')
        $displayIconPath = Resolve-DisplayIconPath -DisplayIcon $item.DisplayIcon
        if ([string]::IsNullOrWhiteSpace($displayIconPath)) {
            $displayIconPath = Resolve-InstallLocationIconPath -InstallLocation $installLocation -DisplayName $displayName
        }
        [PSCustomObject]@{
            名称       = $displayName
            版本       = if ($null -ne $item.DisplayVersion) { ([string]$item.DisplayVersion).Trim() } else { '' }
            发布者     = if ($null -ne $item.Publisher) { ([string]$item.Publisher).Trim() } else { '' }
            安装日期   = if ($null -ne $item.InstallDate) { [string]$item.InstallDate } else { '' }
            安装路径   = $installLocation
            卸载命令   = if ($null -ne $item.UninstallString) { [string]$item.UninstallString } else { '' }
            图标路径   = $displayIconPath
            注册表来源 = ('{0}\{1}\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{2}' -f $item.RegistryHive, $item.RegistryView, $item.RegistryKeyName)
        }
    }

    $deduplicated = foreach ($group in @($result | Group-Object {
        '{0}|{1}|{2}|{3}|{4}|{5}' -f $_.名称, $_.版本, $_.发布者, $_.安装路径, $_.卸载命令, ($_.注册表来源 -replace '\\Registry(?:32|64)\\', '\RegistryView\')
    })) {
        $first = $group.Group | Select-Object -First 1
        if ($group.Count -gt 1) {
            $first.注册表来源 = (($group.Group | Select-Object -ExpandProperty 注册表来源 -Unique) -join '；')
        }
        $first
    }
    return @($deduplicated | Sort-Object 名称, 版本, 安装路径, 注册表来源)
}
