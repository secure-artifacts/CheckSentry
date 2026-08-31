<#
.SYNOPSIS
    采集主流 Chromium 浏览器和 Firefox 插件。
    对同一插件的多个浏览器/个人资料进行聚合，并容错处理损坏的 manifest。
#>

function Get-ExtensionUserProfiles {
    param([switch]$AllUsers)

    $profiles = @()
    $currentProfile = [Environment]::GetFolderPath('UserProfile')
    $currentLocal = [Environment]::GetFolderPath('LocalApplicationData')
    $currentRoaming = [Environment]::GetFolderPath('ApplicationData')
    if (-not [string]::IsNullOrWhiteSpace($currentProfile)) {
        $profiles += [PSCustomObject]@{
            UserName = Split-Path $currentProfile -Leaf
            ProfilePath = $currentProfile
            LocalAppData = $currentLocal
            RoamingAppData = $currentRoaming
        }
    }

    if ($AllUsers) {
        $usersRoot = if (-not [string]::IsNullOrWhiteSpace($env:SystemDrive)) { Join-Path $env:SystemDrive 'Users' } else { '' }
        if (-not [string]::IsNullOrWhiteSpace($usersRoot) -and (Test-Path -LiteralPath $usersRoot -PathType Container)) {
            try {
                foreach ($directory in @(Get-ChildItem -LiteralPath $usersRoot -Directory -Force -ErrorAction Stop)) {
                    $localAppData = Join-Path $directory.FullName 'AppData\Local'
                    $roamingAppData = Join-Path $directory.FullName 'AppData\Roaming'
                    if (-not (Test-Path -LiteralPath $localAppData -PathType Container) -and -not (Test-Path -LiteralPath $roamingAppData -PathType Container)) { continue }
                    if (@($profiles | Where-Object { $_.ProfilePath -eq $directory.FullName }).Count -gt 0) { continue }
                    $profiles += [PSCustomObject]@{
                        UserName = $directory.Name
                        ProfilePath = $directory.FullName
                        LocalAppData = $localAppData
                        RoamingAppData = $roamingAppData
                    }
                }
            } catch {
                Write-Verbose "枚举 Windows 用户目录失败：$($_.Exception.Message)"
            }
        }
    }
    return @($profiles)
}

function Read-BrowserJsonDocument {
    param([string]$Path, [int]$MaxAttempts = 3)
    $lastError = $null
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        $stream = $null
        $reader = $null
        try {
            $shareMode = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
            $stream = New-Object System.IO.FileStream($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $shareMode)
            $reader = New-Object System.IO.StreamReader($stream, (New-Object System.Text.UTF8Encoding($false)), $true)
            $jsonText = $reader.ReadToEnd()
            if ([string]::IsNullOrWhiteSpace($jsonText)) { throw 'JSON 文件内容为空。' }
            return ($jsonText | ConvertFrom-Json -ErrorAction Stop)
        } catch {
            $lastError = $_
            if ($attempt -lt $MaxAttempts) { Start-Sleep -Milliseconds (60 * $attempt) }
        } finally {
            if ($null -ne $reader) { $reader.Dispose() }
            elseif ($null -ne $stream) { $stream.Dispose() }
        }
    }
    if ($null -ne $lastError) { throw $lastError }
    throw "无法读取 JSON 文件：$Path"
}

function Get-SafeChildPath {
    param([string]$BasePath, [string]$RelativePath)
    if ([string]::IsNullOrWhiteSpace($BasePath) -or [string]::IsNullOrWhiteSpace($RelativePath)) { return '' }
    try {
        $baseFull = [System.IO.Path]::GetFullPath($BasePath).TrimEnd('\', '/') + [System.IO.Path]::DirectorySeparatorChar
        $relative = $RelativePath -replace '/', [System.IO.Path]::DirectorySeparatorChar
        $candidate = [System.IO.Path]::GetFullPath((Join-Path $BasePath $relative))
        if (-not $candidate.StartsWith($baseFull, [System.StringComparison]::OrdinalIgnoreCase)) { return '' }
        return $candidate
    } catch {
        return ''
    }
}

function Resolve-ExtensionIconPath {
    param($Manifest, [string]$BasePath)
    try {
        $iconCandidates = @()
        if ($null -ne $Manifest.icons) {
            foreach ($property in $Manifest.icons.PSObject.Properties) {
                $size = 0
                $sizeText = $property.Name -replace '[^0-9]', ''
                if ($sizeText -ne '') { [int]::TryParse($sizeText, [ref]$size) | Out-Null }
                $iconCandidates += [PSCustomObject]@{ Size = $size; RelativePath = [string]$property.Value }
            }
        }
        if ($null -ne $Manifest.icon) {
            $iconCandidates += [PSCustomObject]@{ Size = 0; RelativePath = [string]$Manifest.icon }
        }
        foreach ($candidate in @($iconCandidates | Sort-Object Size -Descending)) {
            $candidatePath = Get-SafeChildPath -BasePath $BasePath -RelativePath $candidate.RelativePath
            if (-not [string]::IsNullOrWhiteSpace($candidatePath) -and (Test-Path -LiteralPath $candidatePath -PathType Leaf)) {
                $item = Get-Item -LiteralPath $candidatePath -ErrorAction SilentlyContinue
                if ($null -ne $item) { return $item.FullName }
            }
        }
    } catch {
        Write-Verbose "读取 Chromium 插件图标失败：$($_.Exception.Message)"
    }
    return ''
}

function Resolve-FirefoxIconPath {
    param($Addon, [string]$ProfilePath)
    try {
        $candidates = @()
        if ($Addon.iconURL) { $candidates += [string]$Addon.iconURL }
        if ($Addon.icons) {
            foreach ($property in $Addon.icons.PSObject.Properties) { $candidates += [string]$property.Value }
        }
        foreach ($candidateValue in $candidates) {
            if ([string]::IsNullOrWhiteSpace($candidateValue)) { continue }
            $candidate = $candidateValue
            if ($candidate -match '^file:///(.*)$') {
                $candidate = [System.Uri]::UnescapeDataString($matches[1]) -replace '/', '\'
                if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                    $item = Get-Item -LiteralPath $candidate -ErrorAction SilentlyContinue
                    if ($null -ne $item) { return $item.FullName }
                }
                continue
            }
            if ($candidate -match '^(jar:|resource:|chrome:|moz-extension:)') { continue }
            if ([System.IO.Path]::IsPathRooted($candidate)) {
                if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                    $item = Get-Item -LiteralPath $candidate -ErrorAction SilentlyContinue
                    if ($null -ne $item) { return $item.FullName }
                }
                continue
            }
            $candidatePath = Get-SafeChildPath -BasePath $ProfilePath -RelativePath $candidate
            if (-not [string]::IsNullOrWhiteSpace($candidatePath) -and (Test-Path -LiteralPath $candidatePath -PathType Leaf)) {
                $item = Get-Item -LiteralPath $candidatePath -ErrorAction SilentlyContinue
                if ($null -ne $item) { return $item.FullName }
            }
        }
    } catch {
        Write-Verbose "读取 Firefox 插件图标失败：$($_.Exception.Message)"
    }
    return ''
}

function Get-ExtensionVersionKey {
    param([object]$VersionText)
    $text = [string]$VersionText
    $parts = @([regex]::Matches($text, '\d+') | ForEach-Object { $_.Value })
    $keyParts = @()
    for ($index = 0; $index -lt 8; $index++) {
        $number = [long]0
        if ($index -lt $parts.Count) { [long]::TryParse($parts[$index], [ref]$number) | Out-Null }
        $keyParts += ('{0:D12}' -f $number)
    }
    return (($keyParts -join '.') + '|' + $text.ToLowerInvariant())
}

function Get-NormalizedPublisher {
    param($Manifest)
    foreach ($candidate in @($Manifest.author, $Manifest.developer)) {
        if ($null -eq $candidate) { continue }
        if ($candidate -is [string]) {
            if (-not [string]::IsNullOrWhiteSpace($candidate)) { return $candidate.Trim() }
            continue
        }
        foreach ($propertyName in @('name', 'email', 'url')) {
            $property = $candidate.PSObject.Properties[$propertyName]
            if ($null -ne $property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
                return ([string]$property.Value).Trim()
            }
        }
    }
    return ''
}

function Resolve-ManifestMessage {
    param($Manifest, [string]$BasePath, [string]$MessageKey)
    if ([string]::IsNullOrWhiteSpace($MessageKey)) { return '' }
    $localesPath = Join-Path $BasePath '_locales'
    if (-not (Test-Path -LiteralPath $localesPath -PathType Container)) { return '' }

    $localeNames = @()
    if ($Manifest.default_locale) { $localeNames += [string]$Manifest.default_locale }
    try {
        $culture = [System.Globalization.CultureInfo]::CurrentUICulture
        $localeNames += $culture.Name.Replace('-', '_')
        $localeNames += $culture.TwoLetterISOLanguageName
    } catch {
        Write-Verbose "读取当前界面语言失败：$($_.Exception.Message)"
    }
    $localeNames += @('en', 'en_US', 'zh_CN', 'zh_TW')
    try {
        $localeNames += @(Get-ChildItem -LiteralPath $localesPath -Directory -ErrorAction Stop | Select-Object -ExpandProperty Name)
    } catch {
        Write-Verbose "枚举插件语言目录失败：$($_.Exception.Message)"
    }

    foreach ($localeName in @($localeNames | Where-Object { $_ -match '^[A-Za-z0-9_-]+$' } | Select-Object -Unique)) {
        $messagePath = Join-Path (Join-Path $localesPath $localeName) 'messages.json'
        if (-not (Test-Path -LiteralPath $messagePath -PathType Leaf)) { continue }
        try {
            $messages = Read-BrowserJsonDocument -Path $messagePath
            $messageProperty = $messages.PSObject.Properties[$MessageKey]
            if ($null -ne $messageProperty -and $null -ne $messageProperty.Value.message) {
                $resolved = [string]$messageProperty.Value.message
                if (-not [string]::IsNullOrWhiteSpace($resolved)) { return $resolved }
            }
        } catch {
            Write-Verbose "解析插件语言文件失败：$messagePath。$($_.Exception.Message)"
        }
    }
    return ''
}

function Get-ChromiumProfileDisplayName {
    param($LocalState, [string]$DirectoryName)
    try {
        $profileProperty = $LocalState.profile.info_cache.PSObject.Properties[$DirectoryName]
        if ($null -ne $profileProperty -and -not [string]::IsNullOrWhiteSpace([string]$profileProperty.Value.name)) {
            return [string]$profileProperty.Value.name
        }
    } catch {
        Write-Verbose "读取 Chromium 个人资料名称失败：$($_.Exception.Message)"
    }
    return $DirectoryName
}

function Get-ChromiumProfileExtensionSettings {
    param([string]$ProfilePath)
    $settings = @{}
    foreach ($fileName in @('Preferences', 'Secure Preferences')) {
        $settingsPath = Join-Path $ProfilePath $fileName
        if (-not (Test-Path -LiteralPath $settingsPath -PathType Leaf)) { continue }
        try {
            $document = Read-BrowserJsonDocument -Path $settingsPath
            $extensionSettings = $document.extensions.settings
            if ($null -eq $extensionSettings) { continue }
            foreach ($property in $extensionSettings.PSObject.Properties) {
                $settings[$property.Name.ToLowerInvariant()] = $property.Value
            }
        } catch {
            Write-Verbose "解析 Chromium 插件启用状态失败：$settingsPath。$($_.Exception.Message)"
        }
    }
    return $settings
}

function Resolve-ChromiumConfiguredExtensionPath {
    param([string]$ProfilePath, [string]$ExtensionsPath, [object]$ConfiguredPath)
    $pathText = ([string]$ConfiguredPath).Trim().Trim('"')
    if ([string]::IsNullOrWhiteSpace($pathText)) { return '' }

    $candidates = @()
    try {
        $expandedPath = [Environment]::ExpandEnvironmentVariables($pathText)
        if ([System.IO.Path]::IsPathRooted($expandedPath)) {
            $candidates += [System.IO.Path]::GetFullPath($expandedPath)
        } else {
            $fromExtensions = Get-SafeChildPath -BasePath $ExtensionsPath -RelativePath $expandedPath
            $fromProfile = Get-SafeChildPath -BasePath $ProfilePath -RelativePath $expandedPath
            if (-not [string]::IsNullOrWhiteSpace($fromExtensions)) { $candidates += $fromExtensions }
            if (-not [string]::IsNullOrWhiteSpace($fromProfile)) { $candidates += $fromProfile }
        }
    } catch {
        return ''
    }

    foreach ($candidateValue in @($candidates | Select-Object -Unique)) {
        $candidate = [string]$candidateValue
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            if ((Split-Path $candidate -Leaf) -ne 'manifest.json') { continue }
            $candidate = Split-Path $candidate -Parent
        }
        if (-not (Test-Path -LiteralPath $candidate -PathType Container)) { continue }
        if (Test-Path -LiteralPath (Join-Path $candidate 'manifest.json') -PathType Leaf) {
            try { return (Get-Item -LiteralPath $candidate -ErrorAction Stop).FullName } catch { continue }
        }
    }
    return ''
}

function Test-IsChromiumComponentSetting {
    param($ExtensionSetting)
    if ($null -eq $ExtensionSetting -or $null -eq $ExtensionSetting.location) { return $false }
    $location = 0
    if (-not [int]::TryParse([string]$ExtensionSetting.location, [ref]$location)) { return $false }
    return $location -in @(5, 10)
}

function Get-ChromiumExtensions {
    param([switch]$IncludeInactive, [switch]$AllUsers)

    $browserRoots = @()
    foreach ($userProfile in @(Get-ExtensionUserProfiles -AllUsers:$AllUsers)) {
        $localDefinitions = [ordered]@{
            'Chrome'   = 'Google\Chrome\User Data'
            'Edge'     = 'Microsoft\Edge\User Data'
            'Chromium' = 'Chromium\User Data'
            'Brave'    = 'BraveSoftware\Brave-Browser\User Data'
            'Vivaldi'  = 'Vivaldi\User Data'
            'Whale'    = 'Naver\Naver Whale\User Data'
        }
        foreach ($browserName in $localDefinitions.Keys) {
            $browserRoots += [PSCustomObject]@{
                BrowserName = $browserName
                UserName = $userProfile.UserName
                Path = Join-Path $userProfile.LocalAppData $localDefinitions[$browserName]
            }
        }
    }

    $allExtensions = @()
    foreach ($browserRoot in $browserRoots) {
        $browserName = $browserRoot.BrowserName
        $userDataPath = $browserRoot.Path
        if (-not (Test-Path -LiteralPath $userDataPath -PathType Container)) { continue }

        $localState = $null
        $localStatePath = Join-Path $userDataPath 'Local State'
        if (Test-Path -LiteralPath $localStatePath -PathType Leaf) {
            try {
                $localState = Read-BrowserJsonDocument -Path $localStatePath
            } catch {
                Write-Verbose "解析 Chromium Local State 失败：$localStatePath。$($_.Exception.Message)"
            }
        }

        try {
            $profiles = @()
            if (Test-Path -LiteralPath (Join-Path $userDataPath 'Extensions') -PathType Container) {
                $profiles += Get-Item -LiteralPath $userDataPath -ErrorAction Stop
            }
            $profiles += @(Get-ChildItem -LiteralPath $userDataPath -Directory -ErrorAction Stop | Where-Object {
                (Test-Path -LiteralPath (Join-Path $_.FullName 'Extensions') -PathType Container) -or
                (Test-Path -LiteralPath (Join-Path $_.FullName 'Preferences') -PathType Leaf) -or
                (Test-Path -LiteralPath (Join-Path $_.FullName 'Secure Preferences') -PathType Leaf)
            })
        } catch {
            continue
        }

        foreach ($browserProfile in $profiles) {
            $extensionsPath = Join-Path $browserProfile.FullName 'Extensions'
            $configuredExtensions = Get-ChromiumProfileExtensionSettings -ProfilePath $browserProfile.FullName
            try {
                $extensionCandidates = @(Get-ChildItem -LiteralPath $extensionsPath -Directory -ErrorAction Stop | ForEach-Object {
                    [PSCustomObject]@{ Name = $_.Name; FullName = $_.FullName; DirectVersionPath = '' }
                })
            } catch { $extensionCandidates = @() }

            $standardPaths = @{}
            foreach ($candidate in $extensionCandidates) { $standardPaths[$candidate.Name.ToLowerInvariant()] = $candidate.FullName }
            foreach ($configuredEntry in $configuredExtensions.GetEnumerator()) {
                $configuredId = ([string]$configuredEntry.Key).ToLowerInvariant()
                if (Test-IsChromiumComponentSetting -ExtensionSetting $configuredEntry.Value) { continue }
                $directPath = Resolve-ChromiumConfiguredExtensionPath -ProfilePath $browserProfile.FullName -ExtensionsPath $extensionsPath -ConfiguredPath $configuredEntry.Value.path
                if ([string]::IsNullOrWhiteSpace($directPath)) { continue }
                if ($standardPaths.ContainsKey($configuredId)) {
                    $standardRoot = ([string]$standardPaths[$configuredId]).TrimEnd('\', '/') + [System.IO.Path]::DirectorySeparatorChar
                    if ($directPath.StartsWith($standardRoot, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
                    $extensionCandidates = @($extensionCandidates | Where-Object { $_.Name -ne $configuredId })
                }
                $extensionCandidates += [PSCustomObject]@{ Name = $configuredId; FullName = $directPath; DirectVersionPath = $directPath }
            }

            foreach ($extensionCandidate in $extensionCandidates) {
                $extensionId = $extensionCandidate.Name
                if ($extensionId -notmatch '^[A-Za-z0-9_@.{}-]{1,256}$') { continue }

                $extensionSettings = $null
                $settingsKey = $extensionId.ToLowerInvariant()
                if ($configuredExtensions.ContainsKey($settingsKey)) { $extensionSettings = $configuredExtensions[$settingsKey] }
                if (Test-IsChromiumComponentSetting -ExtensionSetting $extensionSettings) { continue }
                $active = $true
                if ($null -ne $extensionSettings -and $null -ne $extensionSettings.state) { $active = ([int]$extensionSettings.state -ne 0) }
                if (-not $IncludeInactive -and -not $active) { continue }
                if ($configuredExtensions.Count -gt 0 -and $null -eq $extensionSettings) {
                    # 扩展目录存在但浏览器配置中没有记录时，通常是卸载后的残留目录。
                    continue
                }

                if (-not [string]::IsNullOrWhiteSpace([string]$extensionCandidate.DirectVersionPath)) {
                    $versionDirectories = @([PSCustomObject]@{ Name = Split-Path $extensionCandidate.DirectVersionPath -Leaf; FullName = $extensionCandidate.DirectVersionPath })
                } else {
                    try {
                        $versionDirectories = @(Get-ChildItem -LiteralPath $extensionCandidate.FullName -Directory -ErrorAction Stop | Sort-Object {
                            Get-ExtensionVersionKey -VersionText $_.Name
                        } -Descending)
                    } catch {
                        continue
                    }
                }

                if ($null -ne $extensionSettings -and -not [string]::IsNullOrWhiteSpace([string]$extensionSettings.path)) {
                    $configuredVersionDirectory = Split-Path ([string]$extensionSettings.path).TrimEnd('\', '/') -Leaf
                    $preferred = @($versionDirectories | Where-Object { $_.Name -eq $configuredVersionDirectory })
                    if ($preferred.Count -gt 0) {
                        $versionDirectories = @($preferred + @($versionDirectories | Where-Object { $_.Name -ne $configuredVersionDirectory }))
                    }
                }

                $selected = $null
                foreach ($versionDirectory in $versionDirectories) {
                    $manifestPath = Join-Path $versionDirectory.FullName 'manifest.json'
                    if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) { continue }
                    try {
                        $manifest = Read-BrowserJsonDocument -Path $manifestPath
                        $selected = [PSCustomObject]@{ Directory = $versionDirectory; Manifest = $manifest }
                        break
                    } catch {
                        Write-Verbose "解析插件 manifest 失败：$manifestPath。$($_.Exception.Message)"
                    }
                }
                if ($null -eq $selected) { continue }

                $manifest = $selected.Manifest
                if ($null -ne $manifest.theme -or $null -ne $manifest.app) { continue }
                $name = [string]$manifest.name
                if ($name -match '^__MSG_(.+)__$') {
                    $resolvedName = Resolve-ManifestMessage -Manifest $manifest -BasePath $selected.Directory.FullName -MessageKey $matches[1]
                    if ([string]::IsNullOrWhiteSpace($resolvedName)) { $name = $extensionId } else { $name = $resolvedName }
                }
                if ([string]::IsNullOrWhiteSpace($name)) { $name = $extensionId }

                $allExtensions += [PSCustomObject]@{
                    ExtensionId   = $extensionId
                    Name          = $name
                    Version       = [string]$manifest.version
                    VersionKey    = Get-ExtensionVersionKey -VersionText $manifest.version
                    Publisher     = Get-NormalizedPublisher -Manifest $manifest
                    BrowserFamily = $browserName
                    WindowsUser   = $browserRoot.UserName
                    ProfileName   = Get-ChromiumProfileDisplayName -LocalState $localState -DirectoryName $browserProfile.Name
                    ProfilePath   = $browserProfile.FullName
                    IconPath      = Resolve-ExtensionIconPath -Manifest $manifest -BasePath $selected.Directory.FullName
                    Active        = $active
                }
            }
        }
    }

    $results = @()
    foreach ($group in @($allExtensions | Group-Object ExtensionId)) {
        $items = @($group.Group | Sort-Object VersionKey -Descending)
        if ($items.Count -eq 0) { continue }
        $first = $items[0]
        $locations = foreach ($item in $items) {
            [PSCustomObject]@{
                Browser = $item.BrowserFamily
                WindowsUser = $item.WindowsUser
                ProfileName = $item.ProfileName
                ProfilePath = $item.ProfilePath
                Version = $item.Version
            }
        }
        $iconItem = $items | Where-Object { -not [string]::IsNullOrWhiteSpace($_.IconPath) } | Select-Object -First 1
        $uniqueVersions = @($items | Select-Object -ExpandProperty Version -Unique)
        $displayVersion = if ($uniqueVersions.Count -gt 1) { $uniqueVersions -join ', ' } else { $first.Version }
        $results += [PSCustomObject]@{
            ExtensionId = $group.Name
            Name = $first.Name
            Version = $displayVersion
            Versions = $uniqueVersions
            Publisher = $first.Publisher
            IconPath = if ($null -ne $iconItem) { $iconItem.IconPath } else { '' }
            BrowserFamily = @($items | Select-Object -ExpandProperty BrowserFamily -Unique)
            Locations = @($locations | Sort-Object WindowsUser, Browser, ProfilePath -Unique)
            Ecosystem = 'Chromium'
            Active = (@($items | Where-Object { $_.Active }).Count -gt 0)
        }
    }
    return $results
}

function Get-FirefoxProfileRecords {
    param([string]$ProfilesIniPath, [string]$FirefoxRoot)
    if (-not (Test-Path -LiteralPath $ProfilesIniPath -PathType Leaf)) { return @() }
    $records = @()
    $sectionName = ''
    $section = @{}
    $lines = Get-Content -LiteralPath $ProfilesIniPath -ErrorAction Stop
    foreach ($lineValue in @($lines + '[__END__]')) {
        $line = [string]$lineValue
        if ($line -match '^\s*\[([^]]+)\]\s*$') {
            if ($sectionName -match '^Profile' -and $section.ContainsKey('Path')) {
                $isRelative = -not $section.ContainsKey('IsRelative') -or [string]$section.IsRelative -eq '1'
                $profilePath = if ($isRelative) { Join-Path $FirefoxRoot ([string]$section.Path) } else { [string]$section.Path }
                $records += [PSCustomObject]@{
                    Name = if ($section.ContainsKey('Name')) { [string]$section.Name } else { Split-Path $profilePath -Leaf }
                    Path = $profilePath
                }
            }
            $sectionName = $matches[1]
            $section = @{}
            continue
        }
        if ($line -match '^\s*([^=;#]+?)\s*=\s*(.*)$') { $section[$matches[1]] = $matches[2] }
    }
    return $records
}

function Test-IsFirefoxSystemAddon {
    param($Addon)
    if ($Addon.isSystem -eq $true -or $Addon.isBuiltin -eq $true) { return $true }
    $location = [string]$Addon.location
    return $location -match '^(app-system-defaults|app-system-addons|app-system-share|app-system-local|app-builtin)'
}

function Get-FirefoxExtensions {
    param([switch]$IncludeInactive, [switch]$AllUsers)

    $allExtensions = @()
    foreach ($userProfile in @(Get-ExtensionUserProfiles -AllUsers:$AllUsers)) {
        $firefoxRoot = Join-Path $userProfile.RoamingAppData 'Mozilla\Firefox'
        $profilesIniPath = Join-Path $firefoxRoot 'profiles.ini'
        try { $profiles = @(Get-FirefoxProfileRecords -ProfilesIniPath $profilesIniPath -FirefoxRoot $firefoxRoot) } catch { continue }
        foreach ($firefoxProfile in $profiles) {
        $extensionsPath = Join-Path $firefoxProfile.Path 'extensions.json'
        if (-not (Test-Path -LiteralPath $extensionsPath -PathType Leaf)) { continue }
        try {
            $extensionData = Read-BrowserJsonDocument -Path $extensionsPath
        } catch {
            continue
        }
        foreach ($addon in @($extensionData.addons)) {
            if ($null -eq $addon -or [string]$addon.type -ne 'extension') { continue }
            if (Test-IsFirefoxSystemAddon -Addon $addon) { continue }
            if (-not $IncludeInactive -and $addon.active -ne $true) { continue }
            $extensionId = [string]$addon.id
            if ([string]::IsNullOrWhiteSpace($extensionId)) { continue }
            $name = ''
            if ($null -ne $addon.defaultLocale) { $name = [string]$addon.defaultLocale.name }
            if ([string]::IsNullOrWhiteSpace($name) -and $null -ne $addon.locale) { $name = [string]$addon.locale.name }
            if ([string]::IsNullOrWhiteSpace($name)) { $name = $extensionId }
            $publisher = ''
            if ($null -ne $addon.defaultLocale -and $null -ne $addon.defaultLocale.creator) { $publisher = [string]$addon.defaultLocale.creator }

            $allExtensions += [PSCustomObject]@{
                ExtensionId = $extensionId
                Name = $name
                Version = [string]$addon.version
                VersionKey = Get-ExtensionVersionKey -VersionText $addon.version
                Publisher = $publisher
                IconPath = Resolve-FirefoxIconPath -Addon $addon -ProfilePath $firefoxProfile.Path
                Active = [bool]$addon.active
                SignedState = $addon.signedState
                WindowsUser = $userProfile.UserName
                ProfileName = $firefoxProfile.Name
                ProfilePath = $firefoxProfile.Path
            }
        }
        }
    }

    $results = @()
    foreach ($group in @($allExtensions | Group-Object ExtensionId)) {
        $items = @($group.Group | Sort-Object VersionKey -Descending)
        if ($items.Count -eq 0) { continue }
        $first = $items[0]
        $iconItem = $items | Where-Object { -not [string]::IsNullOrWhiteSpace($_.IconPath) } | Select-Object -First 1
        $locations = foreach ($item in $items) {
            [PSCustomObject]@{ Browser = 'Firefox'; WindowsUser = $item.WindowsUser; ProfileName = $item.ProfileName; ProfilePath = $item.ProfilePath; Version = $item.Version }
        }
        $uniqueVersions = @($items | Select-Object -ExpandProperty Version -Unique)
        $displayVersion = if ($uniqueVersions.Count -gt 1) { $uniqueVersions -join ', ' } else { $first.Version }
        $results += [PSCustomObject]@{
            ExtensionId = $group.Name
            Name = $first.Name
            Version = $displayVersion
            Versions = $uniqueVersions
            Publisher = $first.Publisher
            IconPath = if ($null -ne $iconItem) { $iconItem.IconPath } else { '' }
            Active = (@($items | Where-Object { $_.Active }).Count -gt 0)
            SignedState = $first.SignedState
            BrowserFamily = @('Firefox')
            Locations = @($locations | Sort-Object WindowsUser, ProfilePath -Unique)
            Ecosystem = 'Firefox'
        }
    }
    return $results
}
