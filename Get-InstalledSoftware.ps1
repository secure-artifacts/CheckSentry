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

function New-SoftwareIconReference {
    param(
        [string]$Path,
        [int]$Index,
        [string]$Source
    )
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    return [PSCustomObject]@{ Path = $Path; Index = $Index; Source = $Source }
}

function Resolve-DisplayIconReference {
    param([object]$DisplayIcon)
    $value = Expand-RegistryPathValue -Value $DisplayIcon
    if ([string]::IsNullOrWhiteSpace($value)) { return $null }

    $candidate = $value.Trim()
    if ($candidate.StartsWith('@')) { $candidate = $candidate.Substring(1).Trim() }
    $iconIndex = 0
    if ($candidate -match '^(.*),\s*(-?\d+)\s*$') {
        $candidate = $matches[1].Trim()
        $parsedIndex = 0
        if ([int]::TryParse($matches[2], [ref]$parsedIndex)) { $iconIndex = $parsedIndex }
    }
    $candidate = $candidate.Trim().Trim('"')
    if ($candidate -match '^(.+?\.(?:exe|dll|ico|png|jpg|jpeg|gif))(?:\s+.*)?$') {
        $candidate = $matches[1]
    }

    $candidate = $candidate.Trim().Trim('"')
    if (Test-Path -LiteralPath $candidate -PathType Leaf) {
        $item = Get-Item -LiteralPath $candidate -ErrorAction SilentlyContinue
        if ($null -ne $item) { return New-SoftwareIconReference -Path $item.FullName -Index $iconIndex -Source 'DisplayIcon' }
    }
    Write-Verbose "注册表 DisplayIcon 路径不可用：$value"
    return $null
}

function Resolve-InstallLocationIconReference {
    param(
        [string]$InstallLocation,
        [string]$DisplayName
    )
    if ([string]::IsNullOrWhiteSpace($InstallLocation)) { return $null }
    $base = (Expand-RegistryPathValue -Value $InstallLocation).Trim().Trim('"')
    if (-not (Test-Path -LiteralPath $base -PathType Container)) { return $null }

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
            if ($null -ne $item) { return New-SoftwareIconReference -Path $item.FullName -Index 0 -Source 'InstallLocation' }
        }
    }

    try {
        $normalizedName = ([regex]::Replace($name.ToLowerInvariant(), '[^a-z0-9]+', ''))
        $fallback = Get-ChildItem -LiteralPath $base -Filter '*.exe' -File -ErrorAction Stop |
            Where-Object { $_.Name -notmatch '^(unins|uninstall|setup|update|crash|helper)' } |
            ForEach-Object {
                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
                $normalizedBaseName = [regex]::Replace($baseName.ToLowerInvariant(), '[^a-z0-9]+', '')
                $score = 0
                if (-not [string]::IsNullOrWhiteSpace($normalizedName) -and -not [string]::IsNullOrWhiteSpace($normalizedBaseName) -and $normalizedName.Contains($normalizedBaseName)) { $score = 2 }
                elseif (-not [string]::IsNullOrWhiteSpace($normalizedName) -and -not [string]::IsNullOrWhiteSpace($normalizedBaseName) -and $normalizedBaseName.Contains($normalizedName)) { $score = 1 }
                [PSCustomObject]@{ File = $_; Score = $score }
            } |
            Where-Object { $_.Score -gt 0 } |
            Sort-Object @{ Expression = { $_.Score }; Descending = $true }, @{ Expression = { $_.File.Name }; Descending = $false } |
            Select-Object -First 1
        if ($null -ne $fallback) { return New-SoftwareIconReference -Path $fallback.File.FullName -Index 0 -Source 'InstallLocation' }
    } catch {
        Write-Verbose "枚举安装目录中的图标候选程序失败：$base。$($_.Exception.Message)"
    }
    return $null
}

function Resolve-UninstallStringIconReference {
    param([object]$UninstallString)
    $value = Expand-RegistryPathValue -Value $UninstallString
    if ([string]::IsNullOrWhiteSpace($value)) { return $null }

    $candidate = ''
    if ($value -match '(?i)"([^"]+\.(?:exe|dll))"') {
        $candidate = $matches[1]
    } elseif ($value -match '(?i)([A-Za-z]:\\.*?\.(?:exe|dll))(?=\s|$)') {
        $candidate = $matches[1]
    }
    if ([string]::IsNullOrWhiteSpace($candidate)) { return $null }

    $candidate = $candidate.Trim().Trim('"')
    if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) { return $null }
    $item = Get-Item -LiteralPath $candidate -ErrorAction SilentlyContinue
    if ($null -eq $item) { return $null }
    return New-SoftwareIconReference -Path $item.FullName -Index 0 -Source 'UninstallString'
}

function Resolve-TeamsMeetingAddinIconReference {
    param([string]$DisplayName)
    if ([string]$DisplayName -notlike '*Teams Meeting Add-in*') { return $null }

    $localAppData = [Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)
    if ([string]::IsNullOrWhiteSpace($localAppData)) { return $null }
    $root = Join-Path $localAppData 'Microsoft\TeamsMeetingAdd-in'
    if (-not (Test-Path -LiteralPath $root -PathType Container)) { return $null }

    $versions = @()
    foreach ($directory in @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue)) {
        try {
            $version = [version]$directory.Name
            $versions += [PSCustomObject]@{ Directory = $directory; Version = $version }
        } catch {
            continue
        }
    }

    $architectures = if ([Environment]::Is64BitOperatingSystem) { @('x64', 'x86') } else { @('x86', 'x64') }
    foreach ($versionItem in @($versions | Sort-Object @{ Expression = { $_.Version }; Descending = $true })) {
        foreach ($architecture in $architectures) {
            $candidate = Join-Path $versionItem.Directory.FullName (Join-Path $architecture 'Assets\NewMeeting_Large_192.png')
            if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                $item = Get-Item -LiteralPath $candidate -ErrorAction SilentlyContinue
                if ($null -ne $item) { return New-SoftwareIconReference -Path $item.FullName -Index 0 -Source 'TeamsMeetingAddinAsset' }
            }
        }
    }
    return $null
}

function Initialize-MsiProductInfoApi {
    if ($null -ne ('CheckSentry.MsiProductInfo' -as [type])) { return }
    $source = @'
using System;
using System.Runtime.InteropServices;
using System.Text;

namespace CheckSentry {
    public static class MsiProductInfo {
        [DllImport("msi.dll", CharSet = CharSet.Unicode)]
        public static extern uint MsiGetProductInfo(
            string productCode,
            string property,
            StringBuilder valueBuffer,
            ref int valueLength);
    }
}
'@
    Add-Type -TypeDefinition $source -ErrorAction Stop
}

function Resolve-MsiProductIconReference {
    param(
        [object]$RegistryKeyName,
        [object]$UninstallString
    )
    $productCode = ''
    foreach ($value in @([string]$RegistryKeyName, [string]$UninstallString)) {
        if ($value -match '(?i)\{[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\}') {
            $productCode = $matches[0]
            break
        }
    }
    if ([string]::IsNullOrWhiteSpace($productCode)) { return $null }

    try {
        Initialize-MsiProductInfoApi
        $capacity = 32767
        $buffer = New-Object System.Text.StringBuilder $capacity
        $length = $capacity
        $result = [CheckSentry.MsiProductInfo]::MsiGetProductInfo($productCode, 'ProductIcon', $buffer, [ref]$length)
        if ($result -ne 0 -or $buffer.Length -eq 0) { return $null }
        $reference = Resolve-DisplayIconReference -DisplayIcon $buffer.ToString()
        if ($null -ne $reference) { $reference.Source = 'MSIProductIcon'; return $reference }
    } catch {
        Write-Verbose "读取 MSI ProductIcon 失败：$productCode。$($_.Exception.Message)"
    }
    return $null
}

function ConvertTo-SoftwareLookupKey {
    param([object]$Value)
    $text = ([string]$Value).Trim().ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($text)) { return '' }
    return [regex]::Replace($text, '[^\p{L}\p{Nd}]+', '')
}

function Get-SoftwareNameTokens {
    param([object]$Value)
    return @(([regex]::Split(([string]$Value).ToLowerInvariant(), '[^\p{L}\p{Nd}]+')) | Where-Object {
        $_.Length -ge 2 -and $_ -notmatch '^\d+(?:\.\d+)*$'
    } | Select-Object -Unique)
}

function Get-StartMenuIconCatalog {
    $catalog = @()
    $shell = $null
    try {
        $shell = New-Object -ComObject WScript.Shell -ErrorAction Stop
        $roots = @(
            [Environment]::GetFolderPath([Environment+SpecialFolder]::Programs),
            [Environment]::GetFolderPath([Environment+SpecialFolder]::CommonPrograms)
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and (Test-Path -LiteralPath $_ -PathType Container) } | Select-Object -Unique

        foreach ($root in $roots) {
            foreach ($linkFile in @(Get-ChildItem -LiteralPath $root -Filter '*.lnk' -File -Recurse -ErrorAction SilentlyContinue)) {
                $shortcut = $null
                try {
                    $shortcut = $shell.CreateShortcut($linkFile.FullName)
                    $targetPath = (Expand-RegistryPathValue -Value $shortcut.TargetPath).Trim().Trim('"')
                    if ([string]::IsNullOrWhiteSpace($targetPath) -or -not (Test-Path -LiteralPath $targetPath -PathType Leaf)) { continue }

                    $iconReference = Resolve-DisplayIconReference -DisplayIcon $shortcut.IconLocation
                    if ($null -eq $iconReference) {
                        $targetExtension = [System.IO.Path]::GetExtension($targetPath).ToLowerInvariant()
                        if ($targetExtension -notin @('.exe', '.dll', '.ico', '.png', '.jpg', '.jpeg', '.gif')) { continue }
                        $targetItem = Get-Item -LiteralPath $targetPath -ErrorAction SilentlyContinue
                        if ($null -eq $targetItem) { continue }
                        $iconReference = New-SoftwareIconReference -Path $targetItem.FullName -Index 0 -Source 'StartMenuTarget'
                    } else {
                        $iconReference.Source = 'StartMenuIconLocation'
                    }

                    $lookupValues = @([System.IO.Path]::GetFileNameWithoutExtension($linkFile.Name))
                    try {
                        $versionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($targetPath)
                        $lookupValues += @($versionInfo.ProductName, $versionInfo.FileDescription, $versionInfo.OriginalFilename)
                    } catch {
                        Write-Verbose "读取快捷方式目标版本信息失败：$targetPath。$($_.Exception.Message)"
                    }
                    $lookupValues = @($lookupValues | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Select-Object -Unique)
                    $catalog += [PSCustomObject]@{
                        LinkName = [System.IO.Path]::GetFileNameWithoutExtension($linkFile.Name)
                        TargetPath = $targetPath
                        LookupValues = $lookupValues
                        IconReference = $iconReference
                    }
                } catch {
                    Write-Verbose "解析开始菜单快捷方式失败：$($linkFile.FullName)。$($_.Exception.Message)"
                } finally {
                    if ($null -ne $shortcut -and [Runtime.InteropServices.Marshal]::IsComObject($shortcut)) {
                        try { $null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($shortcut) }
                        catch { Write-Verbose "释放快捷方式 COM 对象失败：$($_.Exception.Message)" }
                    }
                }
            }
        }
    } catch {
        Write-Verbose "无法读取开始菜单快捷方式：$($_.Exception.Message)"
    } finally {
        if ($null -ne $shell -and [Runtime.InteropServices.Marshal]::IsComObject($shell)) {
            try { $null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($shell) }
            catch { Write-Verbose "释放 WScript.Shell COM 对象失败：$($_.Exception.Message)" }
        }
    }
    return $catalog
}

function Get-StartMenuCandidateScore {
    param(
        [string]$DisplayName,
        $Candidate
    )
    $displayKey = ConvertTo-SoftwareLookupKey -Value $DisplayName
    if ([string]::IsNullOrWhiteSpace($displayKey)) { return 0 }
    $bestScore = 0
    foreach ($lookupValue in @($Candidate.LookupValues)) {
        $lookupKey = ConvertTo-SoftwareLookupKey -Value $lookupValue
        if ([string]::IsNullOrWhiteSpace($lookupKey)) { continue }
        if ($lookupKey -eq $displayKey) { return 100 }
        if ($lookupKey.Length -ge 5 -and ($displayKey.Contains($lookupKey) -or $lookupKey.Contains($displayKey))) {
            if ($bestScore -lt 85) { $bestScore = 85 }
        }
    }

    $displayTokens = @(Get-SoftwareNameTokens -Value $DisplayName)
    $candidateTokens = @()
    foreach ($lookupValue in @($Candidate.LookupValues)) { $candidateTokens += @(Get-SoftwareNameTokens -Value $lookupValue) }
    $candidateTokens = @($candidateTokens | Select-Object -Unique)
    $matchedTokens = @($displayTokens | Where-Object { $candidateTokens -contains $_ })
    if ($matchedTokens.Count -ge 2) {
        $tokenScore = 60 + [Math]::Min(20, ($matchedTokens.Count * 5))
        if ($tokenScore -gt $bestScore) { $bestScore = $tokenScore }
    } elseif ($displayTokens.Count -eq 1 -and $matchedTokens.Count -eq 1 -and $matchedTokens[0].Length -ge 5) {
        if ($bestScore -lt 60) { $bestScore = 60 }
    }
    return $bestScore
}

function Resolve-StartMenuIconReference {
    param(
        [string]$DisplayName,
        [object[]]$Catalog
    )
    $ranked = @($Catalog | ForEach-Object {
        [PSCustomObject]@{ Candidate = $_; Score = Get-StartMenuCandidateScore -DisplayName $DisplayName -Candidate $_ }
    } | Where-Object { $_.Score -ge 60 } | Sort-Object @{ Expression = { $_.Score }; Descending = $true }, @{ Expression = { $_.Candidate.TargetPath }; Descending = $false })
    if ($ranked.Count -eq 0) { return $null }
    if ($ranked.Count -gt 1 -and $ranked[0].Score -eq $ranked[1].Score -and
        -not [string]::Equals($ranked[0].Candidate.TargetPath, $ranked[1].Candidate.TargetPath, [StringComparison]::OrdinalIgnoreCase)) {
        Write-Verbose "开始菜单存在分数相同的多个图标候选，已放弃自动匹配：$DisplayName"
        return $null
    }
    return $ranked[0].Candidate.IconReference
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

function Get-AppxSoftwareList {
    $items = @()
    try {
        $appNames = @{}
        $shell = $null
        try {
            $shell = New-Object -ComObject WScript.Shell
            # We use Shell.Application for AppsFolder
            $shellApp = New-Object -ComObject Shell.Application
            $appsFolder = $shellApp.Namespace('shell:AppsFolder')
            if ($null -ne $appsFolder) {
                foreach ($item in $appsFolder.Items()) {
                    if ($item.Path -match '^([^!]+)!') {
                        $pfn = $matches[1]
                        if (-not $appNames.ContainsKey($pfn)) {
                            $appNames[$pfn] = $item.Name
                        }
                    }
                }
            }
        } catch {
            Write-Verbose "读取 shell:AppsFolder 失败: $($_.Exception.Message)"
        } finally {
            if ($null -ne $shellApp -and [Runtime.InteropServices.Marshal]::IsComObject($shellApp)) {
                try { $null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($shellApp) } catch {}
            }
            if ($null -ne $shell -and [Runtime.InteropServices.Marshal]::IsComObject($shell)) {
                try { $null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($shell) } catch {}
            }
        }

        $packages = Get-AppxPackage -ErrorAction SilentlyContinue | Where-Object {
            -not $_.IsFramework -and
            -not $_.IsResourcePackage -and
            -not $_.NonRemovable -and
            $_.SignatureKind -ne 'System'
        }

        foreach ($pkg in $packages) {
            $displayName = $pkg.Name
            if ($appNames.ContainsKey($pkg.PackageFamilyName)) {
                $displayName = $appNames[$pkg.PackageFamilyName]
            }

            $iconPath = ''
            $manifestPath = Join-Path $pkg.InstallLocation "AppxManifest.xml"
            if (Test-Path -LiteralPath $manifestPath -PathType Leaf) {
                try {
                    $xml = [xml](Get-Content -LiteralPath $manifestPath -Raw -ErrorAction Stop)
                    
                    # 声明 XML 命名空间，处理可能带命名空间前缀的节点
                    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
                    $ns.AddNamespace("ns", $xml.DocumentElement.NamespaceURI)
                    $ns.AddNamespace("uap", "http://schemas.microsoft.com/appx/manifest/uap/windows10")

                    $logoNode = $xml.SelectSingleNode("//ns:Properties/ns:Logo", $ns)
                    $logoRel = if ($null -ne $logoNode) { $logoNode.InnerText } else { '' }

                    if ([string]::IsNullOrWhiteSpace($logoRel)) {
                        $appNode = $xml.SelectSingleNode("//ns:Applications/ns:Application[1]/uap:VisualElements", $ns)
                        if ($null -ne $appNode) {
                            $logoRel = $appNode.GetAttribute("Square44x44Logo")
                            if ([string]::IsNullOrWhiteSpace($logoRel)) {
                                $logoRel = $appNode.GetAttribute("Square150x150Logo")
                            }
                        }
                    }

                    if (-not [string]::IsNullOrWhiteSpace($logoRel)) {
                        $logoBase = [System.IO.Path]::GetFileNameWithoutExtension($logoRel)
                        $logoDir = [System.IO.Path]::GetDirectoryName($logoRel)
                        $searchDir = Join-Path $pkg.InstallLocation $logoDir
                        if (Test-Path -LiteralPath $searchDir -PathType Container) {
                            $candidates = Get-ChildItem -LiteralPath $searchDir -Filter "$logoBase*.png" -File -ErrorAction SilentlyContinue
                            if ($candidates.Count -gt 0) {
                                $iconPath = $candidates[0].FullName
                            }
                        }
                    }
                } catch {
                    Write-Verbose "读取 AppxManifest 失败: $($pkg.PackageFullName)"
                }
            }

            $items += [PSCustomObject]@{
                DisplayName     = $displayName
                DisplayVersion  = $pkg.Version
                Publisher       = $pkg.Publisher
                InstallDate     = ''
                InstallLocation = $pkg.InstallLocation
                UninstallString = "Remove-AppxPackage -Package '$($pkg.PackageFullName)'"
                DisplayIcon     = $iconPath
                SystemComponent = 0
                ParentKeyName   = ''
                ReleaseType     = ''
                RegistryHive    = 'AppX'
                RegistryView    = 'Default'
                RegistryKeyName = $pkg.PackageFullName
            }
        }
    } catch {
        Write-Verbose "获取 AppxPackage 失败: $($_.Exception.Message)"
    }
    return $items
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

    $raw += @(Get-AppxSoftwareList)

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

    $shortcutCatalog = $null
    $shortcutCatalogLoaded = $false
    $result = foreach ($item in $raw) {
        $displayName = ([string]$item.DisplayName).Trim()
        $installLocation = (Expand-RegistryPathValue -Value $item.InstallLocation).Trim().Trim('"')
        $iconReference = Resolve-DisplayIconReference -DisplayIcon $item.DisplayIcon
        if ($null -eq $iconReference) {
            $iconReference = Resolve-MsiProductIconReference -RegistryKeyName $item.RegistryKeyName -UninstallString $item.UninstallString
        }
        if ($null -eq $iconReference) {
            $iconReference = Resolve-InstallLocationIconReference -InstallLocation $installLocation -DisplayName $displayName
        }
        if ($null -eq $iconReference) {
            $iconReference = Resolve-UninstallStringIconReference -UninstallString $item.UninstallString
        }
        if ($null -eq $iconReference) {
            $iconReference = Resolve-TeamsMeetingAddinIconReference -DisplayName $displayName
        }
        if ($null -eq $iconReference) {
            if (-not $shortcutCatalogLoaded) {
                $shortcutCatalog = @(Get-StartMenuIconCatalog)
                $shortcutCatalogLoaded = $true
            }
            $iconReference = Resolve-StartMenuIconReference -DisplayName $displayName -Catalog $shortcutCatalog
        }
        [PSCustomObject]@{
            名称       = $displayName
            版本       = if ($null -ne $item.DisplayVersion) { ([string]$item.DisplayVersion).Trim() } else { '' }
            发布者     = if ($null -ne $item.Publisher) { ([string]$item.Publisher).Trim() } else { '' }
            安装日期   = if ($null -ne $item.InstallDate) { [string]$item.InstallDate } else { '' }
            安装路径   = $installLocation
            卸载命令   = if ($null -ne $item.UninstallString) { [string]$item.UninstallString } else { '' }
            图标路径   = if ($null -ne $iconReference) { [string]$iconReference.Path } else { '' }
            图标索引   = if ($null -ne $iconReference) { [int]$iconReference.Index } else { 0 }
            图标来源   = if ($null -ne $iconReference) { [string]$iconReference.Source } else { '' }
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
