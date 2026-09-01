param()

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$failures = New-Object System.Collections.Generic.List[string]
function Assert-True([bool]$Condition, [string]$Message) { if (-not $Condition) { $failures.Add($Message) } }

foreach ($scriptName in @('Start-ComplianceCheck.ps1','Get-InstalledSoftware.ps1','Get-InstalledExtensions.ps1')) {
    $tokens = $null; $errors = $null
    $null = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $root $scriptName), [ref]$tokens, [ref]$errors)
    Assert-True ($errors.Count -eq 0) "$scriptName 存在 PowerShell 语法错误：$($errors -join '; ')"
}

. (Join-Path $root 'Start-ComplianceCheck.ps1') -LibraryOnly
. (Join-Path $root 'Get-InstalledSoftware.ps1')
. (Join-Path $root 'Get-InstalledExtensions.ps1')

Assert-True (Test-VersionMatch -InstalledVersion @('1.2.0','1.3.0') -RuleVersion '1.3.*') '多位置版本应按任一版本匹配。'
Assert-True (-not (Test-VersionMatch -InstalledVersion @('1.2.0','1.3.0') -RuleVersion '2.*')) '不存在的版本不应匹配。'
Assert-True (Test-VersionMatch -InstalledVersion '6.45.2' -RuleVersion '[6.45.3, 6.45.2, 6.45.1]') '云端单元格中的版本列表必须按任一版本匹配。'
Assert-True (-not (Test-VersionMatch -InstalledVersion '6.45.9' -RuleVersion '[6.45.3, 6.45.2, 6.45.1]')) '版本列表中不存在的版本不应匹配。'
Assert-True ((Get-SafeChildPath -BasePath $root -RelativePath '..\outside.txt') -eq '') '安全路径函数必须拒绝目录穿越。'
Assert-True ((ConvertTo-CanonicalAddedTime '46257.1465277778') -eq '2026-08-23 03:31') 'Excel 添加时间序列号必须转换为统一日期时间文本。'
Assert-True ((ConvertTo-CanonicalAddedTime ([DateTime]'2026-08-23T03:31:00')) -eq '2026-08-23 03:31') '日期对象写入本地清单前必须转换为统一文本格式。'

$extensionFixtureRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('CheckSentry-extension-fixture-' + [guid]::NewGuid().ToString('N'))
try {
    $extensionProfile = Join-Path $extensionFixtureRoot 'Default'
    $extensionStore = Join-Path $extensionProfile 'Extensions'
    $unpackedExtension = Join-Path $extensionFixtureRoot 'UnpackedExtension'
    $zippedExtension = Join-Path $extensionProfile 'UnpackedExtensions\fixture_ABC123'
    $null = New-Item -ItemType Directory -Path $extensionStore, $unpackedExtension, $zippedExtension -Force
    [System.IO.File]::WriteAllText((Join-Path $unpackedExtension 'manifest.json'), '{"name":"Fixture","version":"1.0"}', (New-Object System.Text.UTF8Encoding($false)))
    [System.IO.File]::WriteAllText((Join-Path $zippedExtension 'manifest.json'), '{"name":"Zip Fixture","version":"1.0"}', (New-Object System.Text.UTF8Encoding($false)))
    $resolvedUnpackedPath = Resolve-ChromiumConfiguredExtensionPath -ProfilePath $extensionProfile -ExtensionsPath $extensionStore -ConfiguredPath $unpackedExtension
    $resolvedZipPath = Resolve-ChromiumConfiguredExtensionPath -ProfilePath $extensionProfile -ExtensionsPath $extensionStore -ConfiguredPath 'UnpackedExtensions\fixture_ABC123'
    Assert-True ($resolvedUnpackedPath -eq $unpackedExtension) '配置中登记的解压加载插件路径必须能够被识别。'
    Assert-True ($resolvedZipPath -eq $zippedExtension) '开发者模式从 ZIP 安装到 UnpackedExtensions 的插件必须能够被识别。'
    Assert-True (Test-IsChromiumComponentSetting -ExtensionSetting ([PSCustomObject]@{ location = 5 })) 'Chromium 内置组件扩展必须被识别为组件。'
    Assert-True (-not (Test-IsChromiumComponentSetting -ExtensionSetting ([PSCustomObject]@{ location = 4 }))) '开发者模式加载的插件不得被误判为 Chromium 内置组件。'
} finally {
    if (Test-Path -LiteralPath $extensionFixtureRoot -PathType Container) { Remove-Item -LiteralPath $extensionFixtureRoot -Recurse -Force }
}

$videoKitDefaults = Get-DefaultRuleFields -Item ([PSCustomObject]@{ 类型='软件'; 名称='VideoKit 4.4.27'; 版本='4.4.27' })
Assert-True ($videoKitDefaults.MatchType -eq '包含') '名称末尾包含版本的软件应默认使用包含匹配。'
Assert-True ($videoKitDefaults.NamePattern -eq 'VideoKit') '软件规则关键词应自动移除名称末尾的版本号。'
Assert-True ($videoKitDefaults.Version -eq '4.4.27') '复制或批量添加规则时必须自动填写当前版本号。'

$legacyVideoKitRule = [PSCustomObject]@{ 类型='软件'; 匹配方式='精确'; 软件名关键词='VideoKit 4.4.28'; 插件ID=''; 版本号=''; 发布者='' }
$legacyVideoKitMatcher = New-RuleMatcher -Rules @($legacyVideoKitRule) -ItemType '软件'
Assert-True ($null -eq (Find-FirstMatchingRule -Rules $legacyVideoKitMatcher -ItemType '软件' -DisplayName 'VideoKit 4.4.27' -Publisher '' -ExtensionId '' -Version '4.4.27')) '旧规则中内嵌的目标版本不应误判为当前版本匹配。'
Assert-True ($null -ne (Find-FirstMatchingRule -Rules $legacyVideoKitMatcher -ItemType '软件' -DisplayName 'VideoKit 4.4.27' -Publisher '' -ExtensionId '' -Version '4.4.27' -IgnoreVersion)) '旧规则名称中的版本号应被自动识别并触发版本变化。'

$duplicatedVersionVideoKitRules = @(
    [PSCustomObject]@{ 类型='软件'; 匹配方式='精确'; 软件名关键词='VideoKit 4.4.25'; 插件ID=''; 版本号='4.4.25'; 发布者='' },
    [PSCustomObject]@{ 类型='软件'; 匹配方式='精确'; 软件名关键词='VideoKit 4.4.28'; 插件ID=''; 版本号='4.4.28'; 发布者='' }
)
$videoKitResult = @(Get-ComplianceResult -Installed @([PSCustomObject]@{ 名称='VideoKit 4.4.27'; 版本='4.4.27'; 发布者=''; 安装日期=''; 安装路径=''; 图标路径=''; 图标索引=0 }) -BlackRules @() -WhiteRules $duplicatedVersionVideoKitRules -PendingRules @() -ItemType '软件')
Assert-True ($videoKitResult.Count -eq 1 -and $videoKitResult[0].状态 -eq '版本变化') '名称和版本列重复包含旧版本时，VideoKit 仍必须进入版本变化。'
Assert-True ([string]$videoKitResult[0].MatchedRule.'版本号' -eq '4.4.25') '版本变化必须保留实际命中的规则版本用于界面显示。'

$fullWidthRule = [PSCustomObject]@{ 类型='软件'; 匹配方式='精确'; 软件名关键词='B＆O Audio Control'; 插件ID=''; 版本号='1.47.308.0'; 发布者='' }
$fullWidthMatcher = New-RuleMatcher -Rules @($fullWidthRule) -ItemType '软件'
Assert-True ($null -ne (Find-FirstMatchingRule -Rules $fullWidthMatcher -ItemType '软件' -DisplayName 'B&O Audio Control' -Publisher '' -ExtensionId '' -Version '1.47.308.0')) '全角与半角标点差异不应导致云端白名单匹配失败。'
Assert-True ((Get-RuleIdentity -ItemType '软件' -ExtensionId '' -NamePattern 'B＆O Audio Control') -eq (Get-RuleIdentity -ItemType '软件' -ExtensionId '' -NamePattern 'B&O Audio Control')) '云端和本地规则身份也必须统一全角与半角字符。'
$replacementPadding = -join @(1..6 | ForEach-Object { [char]0xFFFD })
$monosnapRule = [PSCustomObject]@{ 类型='软件'; 匹配方式='精确'; 软件名关键词='Monosnap'; 插件ID=''; 版本号='5.3.0'; 发布者='Monosnap Inc.' }
$monosnapInstalled = [PSCustomObject]@{ 名称=('Monosnap' + $replacementPadding); 版本=('5.3.0' + $replacementPadding); 发布者=('Monosnap Inc.' + $replacementPadding); 安装日期=''; 安装路径=''; 图标路径=''; 图标索引=0 }
$monosnapResult = @(Get-ComplianceResult -Installed @($monosnapInstalled) -BlackRules @() -WhiteRules @($monosnapRule) -PendingRules @() -ItemType '软件')
Assert-True ($monosnapResult.Count -eq 1 -and $monosnapResult[0].状态 -eq '已匹配') '注册表文本尾部含 Unicode 替换字符时，Monosnap 仍必须准确命中白名单。'
Assert-True ((ConvertTo-CleanSoftwareText ('Monosnap' + $replacementPadding)) -eq 'Monosnap') '软件采集阶段必须清除 Unicode 替换字符。'
$veracryptIdrixIdentity = Get-RuleIdentity -ItemType '软件' -ExtensionId '' -NamePattern 'VeraCrypt' -Publisher 'IDRIX' -Version '1.25.9'
$veracryptAmCryptoIdentity = Get-RuleIdentity -ItemType '软件' -ExtensionId '' -NamePattern 'VeraCrypt' -Publisher 'AM Crypto' -Version '1.26.24'
Assert-True ($veracryptIdrixIdentity -ne $veracryptAmCryptoIdentity) '同名软件的不同正版发布者和版本必须作为不同规则保留。'
$fourKOpenMediaIdentity = Get-RuleIdentity -ItemType '软件' -ExtensionId '' -NamePattern '4K Video Downloader+' -Publisher 'Open Media LLC' -Version ''
$fourKInterPromoIdentity = Get-RuleIdentity -ItemType '软件' -ExtensionId '' -NamePattern '4K Video Downloader+' -Publisher 'InterPromo GMBH' -Version '26.3.1.389'
Assert-True ($fourKOpenMediaIdentity -ne $fourKInterPromoIdentity) '4K Video Downloader+ 的不同发布者和版本必须作为不同规则保留。'
Assert-True ((Get-RuleIdentity -ItemType '软件' -ExtensionId '' -NamePattern 'VideoKit 4.4.28' -Publisher '' -Version '') -eq (Get-RuleIdentity -ItemType '软件' -ExtensionId '' -NamePattern 'VideoKit' -Publisher '' -Version '4.4.28')) '名称内嵌版本与版本列写法必须得到相同规则身份。'
Assert-True ((Get-RuleIdentity -ItemType '软件' -ExtensionId '' -NamePattern 'Tool' -Publisher 'Vendor' -Version '[2.0, 1.0]') -eq (Get-RuleIdentity -ItemType '软件' -ExtensionId '' -NamePattern 'Tool' -Publisher 'Vendor' -Version '[1.0, 2.0]')) '版本列表顺序不应制造伪重复差异。'
$publisherRule = [PSCustomObject]@{ 类型='软件'; 匹配方式='精确'; 软件名关键词='VeraCrypt'; 插件ID=''; 版本号='1.26.24'; 发布者='IDRIX' }
$publisherMatcher = New-RuleMatcher -Rules @($publisherRule) -ItemType '软件'
Assert-True ($null -eq (Find-FirstMatchingRule -Rules $publisherMatcher -ItemType '软件' -DisplayName 'VeraCrypt' -Publisher 'AM Crypto' -ExtensionId '' -Version '1.26.24')) '发布者不一致时不能直接判定白名单命中。'
Assert-True ($null -ne (Find-FirstMatchingRule -Rules $publisherMatcher -ItemType '软件' -DisplayName 'VeraCrypt' -Publisher 'AM Crypto' -ExtensionId '' -Version '1.26.24' -IgnorePublisher)) '发布者变化必须能被识别并给出明确原因。'
$publisherResult = @(Get-ComplianceResult -Installed @([PSCustomObject]@{ 名称='VeraCrypt'; 版本='1.26.24'; 发布者='AM Crypto'; 安装日期=''; 安装路径=''; 图标路径=''; 图标索引=0 }) -BlackRules @() -WhiteRules @($publisherRule) -PendingRules @() -ItemType '软件')
Assert-True ($publisherResult.Count -eq 1 -and $publisherResult[0].状态 -eq '待定' -and $publisherResult[0].原因 -match '发布者不同') '白名单近似匹配必须在待定区明确显示发布者差异。'
$emptyRuleResult = @(Get-ComplianceResult -Installed @([PSCustomObject]@{ 名称='Unknown Tool'; 版本='1.0'; 发布者=''; 安装日期=''; 安装路径=''; 图标路径=''; 图标索引=0 }) -BlackRules @() -WhiteRules @() -PendingRules @() -ItemType '软件')
Assert-True ($emptyRuleResult.Count -eq 1 -and $emptyRuleResult[0].状态 -eq '待定' -and $null -eq $emptyRuleResult[0].MatchedRule) '未配置任何规则时，扫描结果必须正常进入待定。'

$blackSoftwareRule = [PSCustomObject]@{ 类型='软件'; 匹配方式='精确'; 软件名关键词='abc 1.2.3'; 插件ID=''; 版本号='1.2.3'; 发布者='' }
$blackSoftwareResult = @(Get-ComplianceResult -Installed @([PSCustomObject]@{ 名称='abc 1.2.2'; 版本='1.2.2'; 发布者=''; 安装日期=''; 安装路径=''; 图标路径=''; 图标索引=0 }) -BlackRules @($blackSoftwareRule) -WhiteRules @() -PendingRules @() -ItemType '软件')
Assert-True ($blackSoftwareResult.Count -eq 1 -and $blackSoftwareResult[0].状态 -eq '命中黑名单') '黑名单软件名称匹配后必须忽略版本差异。'
$publisherScopedBlackRule = [PSCustomObject]@{ 类型='软件'; 匹配方式='精确'; 软件名关键词='VeraCrypt'; 插件ID=''; 版本号='1.25.9'; 发布者='IDRIX' }
$publisherScopedBlackResult = @(Get-ComplianceResult -Installed @([PSCustomObject]@{ 名称='VeraCrypt'; 版本='1.26.24'; 发布者='AM Crypto'; 安装日期=''; 安装路径=''; 图标路径=''; 图标索引=0 }) -BlackRules @($publisherScopedBlackRule) -WhiteRules @() -PendingRules @() -ItemType '软件')
Assert-True ($publisherScopedBlackResult.Count -eq 1 -and $publisherScopedBlackResult[0].状态 -eq '待定') '黑名单软件填写发布者后不得误杀同名的其他发布者。'

$clickCleanWhiteRule = [PSCustomObject]@{ 类型='Chromium插件'; 匹配方式='插件ID精确'; 软件名关键词='Click&Clean'; 插件ID='ghgabhipcejejjmhhchfonmamedcbeod'; 版本号='9.8.2.0'; 发布者='' }
$clickCleanBlackRule = [PSCustomObject]@{ 类型='Chromium插件'; 匹配方式='插件ID精确'; 软件名关键词='Click&Clean'; 插件ID='dacknjoogbepndbemlmljdobinliojbk'; 版本号='9.8.2.0'; 发布者='旧发布者文本' }
function New-TestExtension([string]$ExtensionId, [string]$Version = '9.8.2.0') {
    return [PSCustomObject]@{ Name='Click&Clean'; ExtensionId=$ExtensionId; Version=$Version; Versions=@($Version); Publisher=''; Locations=@(); IconPath=''; Enabled=$true }
}
$whiteClickCleanResult = @(Get-ComplianceResult -Installed @(New-TestExtension 'ghgabhipcejejjmhhchfonmamedcbeod') -BlackRules @($clickCleanBlackRule) -WhiteRules @($clickCleanWhiteRule) -PendingRules @() -ItemType 'Chromium插件')
Assert-True ($whiteClickCleanResult.Count -eq 1 -and $whiteClickCleanResult[0].状态 -eq '已匹配') '同名插件必须按准确 ID 命中白名单，不能被另一个黑名单 ID 误杀。'
$blackClickCleanResult = @(Get-ComplianceResult -Installed @(New-TestExtension 'dacknjoogbepndbemlmljdobinliojbk' '10.0.0') -BlackRules @($clickCleanBlackRule) -WhiteRules @($clickCleanWhiteRule) -PendingRules @() -ItemType 'Chromium插件')
Assert-True ($blackClickCleanResult.Count -eq 1 -and $blackClickCleanResult[0].状态 -eq '命中黑名单') '黑名单插件 ID 匹配后必须忽略版本和发布者差异。'
$unknownClickCleanResult = @(Get-ComplianceResult -Installed @(New-TestExtension 'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa') -BlackRules @($clickCleanBlackRule) -WhiteRules @($clickCleanWhiteRule) -PendingRules @() -ItemType 'Chromium插件')
Assert-True ($unknownClickCleanResult.Count -eq 1 -and $unknownClickCleanResult[0].状态 -eq '待定' -and $unknownClickCleanResult[0].原因 -match '名称与黑名单插件相同.*插件 ID 不同') '同名但陌生 ID 的插件必须进入待定并给出风险原因。'

$highlightExtensionId = 'ibocimhdhiiccenajinccijhmgenlnlc'
$highlightRuleWithInvisibleText = [PSCustomObject]@{
    类型='Chromium插件'; 匹配方式='插件ID精确'; 软件名关键词='高亮插件'
    插件ID=($highlightExtensionId + [char]0x200B + [char]0x00A0); 版本号='3.0.0'; 发布者=''
}
$highlightInstalled = [PSCustomObject]@{ Name='高亮插件'; ExtensionId=$highlightExtensionId; Version='3.0.0'; Versions=@('3.0.0'); Publisher=''; Locations=@(); IconPath=''; Active=$true }
$highlightResult = @(Get-ComplianceResult -Installed @($highlightInstalled) -BlackRules @() -WhiteRules @($highlightRuleWithInvisibleText) -PendingRules @() -ItemType 'Chromium插件')
Assert-True ($highlightResult.Count -eq 1 -and $highlightResult[0].状态 -eq '已匹配') '云端插件 ID 含零宽字符或不换行空格时，仍必须按同一 ID 命中白名单。'
Assert-True ((ConvertTo-CanonicalExtensionId -Value ('chrome-extension://' + $highlightExtensionId + '/src/scopes/options/index.html') -ItemType 'Chromium插件') -eq $highlightExtensionId) '从地址栏复制的完整 chrome-extension:// 地址必须提取出准确插件 ID。'
$highlightAliasRule = [PSCustomObject]@{ 类型='Chromium插件'; 匹配方式='插件ID精确'; 软件名关键词='高亮插件'; 插件ID=('ioldaogdolojkkfdmhkpnnnhnbklpobh, ' + $highlightExtensionId); 版本号='3.0.0'; 发布者='' }
$highlightAliasResult = @(Get-ComplianceResult -Installed @($highlightInstalled) -BlackRules @() -WhiteRules @($highlightAliasRule) -PendingRules @() -ItemType 'Chromium插件')
Assert-True ($highlightAliasResult.Count -eq 1 -and $highlightAliasResult[0].状态 -eq '已匹配') '同一插件规则配置多个可信 ID 时，任一准确 ID 都必须命中。'
Assert-True ((Get-RuleIdentity -ItemType 'Chromium插件' -ExtensionId 'ibocimhdhiiccenajinccijhmgenlnlc,ioldaogdolojkkfdmhkpnnnhnbklpobh' -NamePattern '高亮插件') -eq (Get-RuleIdentity -ItemType 'Chromium插件' -ExtensionId 'ioldaogdolojkkfdmhkpnnnhnbklpobh;ibocimhdhiiccenajinccijhmgenlnlc' -NamePattern '高亮插件')) '可信插件 ID 别名顺序不应改变规则身份。'
$differentDeveloperExtension = [PSCustomObject]@{ Name='高亮插件'; ExtensionId='aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa'; Version='3.0.0'; Versions=@('3.0.0'); Publisher=''; Locations=@(); IconPath=''; Active=$true; InstallLocations=@(4) }
$secondHighlightRule = [PSCustomObject]@{ 类型='Chromium插件'; 匹配方式='插件ID精确'; 软件名关键词='高亮插件'; 插件ID='ioldaogdolojkkfdmhkpnnnhnbklpobh'; 版本号='3.0.0'; 发布者='' }
$differentDeveloperResult = @(Get-ComplianceResult -Installed @($differentDeveloperExtension) -BlackRules @() -WhiteRules @($highlightRuleWithInvisibleText, $secondHighlightRule) -PendingRules @() -ItemType 'Chromium插件')
Assert-True ($differentDeveloperResult.Count -eq 1 -and $differentDeveloperResult[0].状态 -eq '待定' -and $differentDeveloperResult[0].PendingKind -eq 'PluginWhitelistIdMismatch' -and $differentDeveloperResult[0].原因 -match '开发者插件ID可能随路径变化') '开发者插件实际 ID 与白名单不同时必须保持待定并使用简短路径派生 ID 提示。'
Assert-True (@($differentDeveloperResult[0].RuleExtensionIds).Count -eq 2) '同名插件的全部白名单规则 ID 都必须进入差异提示。'
$differentDeveloperHtml = Get-RowHtml -Item $differentDeveloperResult[0] -NeedsAction $true
Assert-True ($differentDeveloperHtml -match 'ID不一致' -and $differentDeveloperHtml -match '规则ID：.*ibocimhdhiiccenajinccijhmgenlnlc' -and $differentDeveloperHtml -match 'ioldaogdolojkkfdmhkpnnnhnbklpobh' -and $differentDeveloperHtml -match '当前ID：aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa') '待定区域必须醒目标记 ID 不一致，悬停展示全部规则 ID 和当前 ID。'
$pendingOrderHtml = Build-ReportHtml -Results @($emptyRuleResult[0], $differentDeveloperResult[0]) -Path 'list.xlsx' -CsrfToken 'test' -Nonce 'test'
Assert-True ($pendingOrderHtml.IndexOf('ID不一致') -lt $pendingOrderHtml.IndexOf('Unknown Tool')) '待定区域中同名但 ID 不一致的插件必须排在普通待定项上方。'

Assert-XlsxArchiveComplete -Path (Join-Path $root 'list_template.xlsx')
$cloudStats = Get-CloudRuleStats -AllRules ([PSCustomObject]@{ '白名单'=@(1,2); '待定'=@(3); '黑名单'=@(4,5,6) })
Assert-True ($cloudStats.TotalCount -eq 6 -and $cloudStats.WhiteCount -eq 2 -and $cloudStats.PendingCount -eq 1 -and $cloudStats.BlackCount -eq 3) '云端规则分表数量统计必须准确。'

Import-Module (Join-Path $root 'Modules/ImportExcel/7.8.10/ImportExcel.psd1') -Force -ErrorAction Stop
$cloudFixturePath = Join-Path ([System.IO.Path]::GetTempPath()) ('CheckSentry-cloud-fixture-' + [guid]::NewGuid().ToString('N') + '.xlsx')
try {
    Copy-Item -LiteralPath (Join-Path $root 'list_template.xlsx') -Destination $cloudFixturePath -ErrorAction Stop
    $fixturePackage = Open-ExcelPackage -Path $cloudFixturePath -ErrorAction Stop
    try {
        $null = $fixturePackage.Workbook.Worksheets.Add('其他数据')
        $whiteSheet = $fixturePackage.Workbook.Worksheets['白名单']
        $whiteSheet.Cells[2, 2].Value = '软件'
        $whiteSheet.Cells[2, 4].Value = '精确'
        $whiteSheet.Cells[2, 5].Value = 'Length Test Tool'
        $whiteSheet.Cells[2, 6].Value = (('1' * 2049) -join '')
        $whiteSheet.Cells[3, 2].Value = 'Chromium插件'
        $whiteSheet.Cells[3, 3].Value = 'ibocimhdhiiccenajinccijhmgenlnl'
        $whiteSheet.Cells[3, 4].Value = '插件ID精确'
        $whiteSheet.Cells[3, 5].Value = 'Invalid ID Test'
        $whiteSheet.Cells[4, 2].Value = '软件'
        $whiteSheet.Cells[4, 4].Value = '精确'
        $whiteSheet.Cells[4, 5].Value = 'Valid Cloud Tool'
        $fixturePackage.Save()
    } finally { $fixturePackage.Dispose() }

    $cloudImport = Import-CloudRuleSheets -Path $cloudFixturePath
    $cloudWarnings = @($cloudImport.ValidationWarnings)
    Assert-True (@($cloudImport.'白名单').Count -eq 1 -and [string]$cloudImport.'白名单'[0].'软件名关键词' -eq 'Valid Cloud Tool') '云端存在无效规则时必须跳过坏行并继续导入其他有效规则。'
    Assert-True ($cloudWarnings.Count -eq 2) '云端版本字段和插件 ID 错误必须分别形成校验提示，不能阻断整个 UI。'
    Assert-True (@($cloudWarnings | Where-Object { $_.Sheet -eq '白名单' -and $_.Cell -eq 'F2' -and $_.Message -match '版本号.*2049' }).Count -eq 1) '云端规则校验提示必须指出版本错误的工作表、单元格和字段。'
    Assert-True (@($cloudWarnings | Where-Object { $_.Sheet -eq '白名单' -and $_.Cell -eq 'C3' -and $_.Message -match '32 位' }).Count -eq 1) '不足 32 位的 Chromium 插件 ID 必须在校验提示中指出准确单元格。'
    Assert-True (@($cloudWarnings | Where-Object { $_.Message -match '其他数据' }).Count -eq 0) '云端导入必须忽略待定、白名单、黑名单以外的工作表。'
    Set-CloudSyncRuntimeState -Configured $true -Status '有警告' -UsingCache $false -Message '测试警告' -Stats (Get-CloudRuleStats -AllRules $cloudImport) -RuleHash 'TEST' -SyncedAt '2026-01-01 00:00:00' -ValidationWarnings $cloudWarnings
    $cloudWarningReport = Build-ReportHtml -Results @() -Path 'list.xlsx' -CsrfToken 'test' -Nonce 'test'
    Assert-True ($cloudWarningReport -match '不会阻止软件使用' -and $cloudWarningReport -match '白名单！C3' -and $cloudWarningReport -match 'ibocimhdhiiccenajinccijhmgenlnl') '云端坏行必须在报告表格上方持续显示位置、原因和当前值。'
    Set-CloudSyncRuntimeState -Configured $false -Status '未配置' -UsingCache $false -Message '' -Stats $null -RuleHash '' -SyncedAt ''
} finally {
    if (Test-Path -LiteralPath $cloudFixturePath -PathType Leaf) { Remove-Item -LiteralPath $cloudFixturePath -Force }
}

Assert-True ((Get-RuleIdentity -ItemType 'Chromium插件' -ExtensionId 'ibocimhdhiiccenajinccijhmgenlnlc' -NamePattern '高亮插件') -ne (Get-RuleIdentity -ItemType 'Chromium插件' -ExtensionId 'ibocimhdhiiccenajinccijhmgenlnld' -NamePattern '高亮插件')) '云端插件 ID 修改一个字符后必须生成新身份，以便删除旧规则并写入新规则。'
$ownerTestDirectory = Join-Path ([System.IO.Path]::GetTempPath()) ('CheckSentry-cloud-owner-test-' + [guid]::NewGuid().ToString('N'))
$originalDataDirectory = $script:DataDirectory
try {
    $ownerCacheDirectory = Join-Path $ownerTestDirectory 'CloudCache'
    $null = New-Item -ItemType Directory -Path $ownerCacheDirectory -Force
    $ownerSnapshot = Join-Path $ownerCacheDirectory 'last-good.xlsx'
    Copy-Item -LiteralPath (Join-Path $root 'list_template.xlsx') -Destination $ownerSnapshot -ErrorAction Stop
    $ownerPackage = Open-ExcelPackage -Path $ownerSnapshot -ErrorAction Stop
    try {
        $ownerRule = [PSCustomObject]@{ 类型='Chromium插件'; 插件ID='ibocimhdhiiccenajinccijhmgenlnlc'; 匹配方式='插件ID精确'; 软件名关键词='高亮插件'; 版本号='3.0.0'; 发布者=''; '状态/分类'='允许'; '备注/原因'=''; '备注/原因链接'=''; 添加人='test'; 添加时间='2026-01-01' }
        $null = Add-RuleToPackage -Package $ownerPackage -SheetName '白名单' -Rule $ownerRule
        Close-ExcelPackage -ExcelPackage $ownerPackage -ErrorAction Stop
        $ownerPackage = $null
    } finally {
        if ($null -ne $ownerPackage) { $ownerPackage.Dispose() }
    }
    $script:DataDirectory = $ownerTestDirectory
    $previousCloudIdentities = Get-PreviousCloudIdentitySet -Url 'https://docs.google.com/spreadsheets/d/aaaaaaaaaa/edit'
    $oldCloudIdentity = Get-RuleIdentity -ItemType 'Chromium插件' -ExtensionId 'ibocimhdhiiccenajinccijhmgenlnlc' -NamePattern '高亮插件'
    Assert-True ($previousCloudIdentities.ContainsKey($oldCloudIdentity)) '同步元数据缺失时也必须从上次云端快照恢复旧规则身份。'
} finally {
    $script:DataDirectory = $originalDataDirectory
    if (Test-Path -LiteralPath $ownerTestDirectory -PathType Container) { Remove-Item -LiteralPath $ownerTestDirectory -Recurse -Force }
}

$rules = @(
    [PSCustomObject]@{ 类型='软件'; 匹配方式='包含'; 软件名关键词='Tool'; 插件ID=''; 版本号=''; 发布者='' },
    [PSCustomObject]@{ 类型='软件'; 匹配方式='精确'; 软件名关键词='Tool Pro'; 插件ID=''; 版本号=''; 发布者='' }
)
$matcher = New-RuleMatcher -Rules $rules -ItemType '软件'
$matched = Find-FirstMatchingRule -Rules $matcher -ItemType '软件' -DisplayName 'Tool Pro' -Publisher '' -ExtensionId '' -Version '1.0'
Assert-True ([object]::ReferenceEquals($matched, $rules[0])) '规则索引必须保持原始工作表顺序。'

$script:InventoryCache = $null
$script:InventoryScanError = ''
$script:InventoryScanResults = @{}
$script:InventoryScanJobs = @{}
foreach ($kind in @('Software','Chromium','Firefox')) {
    $job = Start-Job -ArgumentList $kind -ScriptBlock { param($Kind); [PSCustomObject]@{ Kind=$Kind; Items=@([PSCustomObject]@{ Name=$Kind }) } }
    $script:InventoryScanJobs[$kind] = [PSCustomObject]@{ Job=$job; Received=$false }
}
$deadline = (Get-Date).AddSeconds(15)
while ($null -eq $script:InventoryCache -and (Get-Date) -lt $deadline) { $null = Update-InventoryScanState; Start-Sleep -Milliseconds 50 }
Assert-True ($null -ne $script:InventoryCache) '后台扫描结果必须能合并到库存缓存。'
Assert-True (@($script:InventoryCache.Software).Count -eq 1) '软件后台扫描结果数量不正确。'
Stop-InventoryScanJobs

$batch = Get-Content -LiteralPath (Join-Path $root '运行核对工具.bat') -Raw
Assert-True ($batch -match '-ExecutionPolicy\s+Bypass') '双击入口必须使用当前进程级 Bypass。'
Assert-True ($batch -match 'chcp\s+65001') 'CMD 启动入口必须切换到 UTF-8 代码页。'
Assert-True ($batch -match 'where\s+wt\.exe' -and $batch -match 'start\s+"CheckSentry"\s+wt\.exe') '双击入口必须优先使用支持中文字体回退的 Windows Terminal。'
$launcherSource = Get-Content -LiteralPath (Join-Path $root 'launcher/Program.cs') -Raw -Encoding UTF8
Assert-True ($launcherSource -match 'SetConsoleOutputCP\(65001\)') 'EXE 启动器必须把控制台切换到 UTF-8。'
Assert-True ($launcherSource -match 'RedirectStandardOutput\s*=\s*false' -and $launcherSource -match 'RedirectStandardError\s*=\s*false') 'PowerShell 必须直接写入控制台，禁止经过会把中文替换为问号的重定向管道。'
Assert-True ($launcherSource -match '"-LogPath",\s*logPath') '启动失败日志必须由 PowerShell 直接按 UTF-8 写入。'
$mainScriptSource = Get-Content -LiteralPath (Join-Path $root 'Start-ComplianceCheck.ps1') -Raw -Encoding UTF8
Assert-True ($mainScriptSource -match '\[Console\]::OutputEncoding\s*=\s*\$script:Utf8ConsoleEncoding' -and $mainScriptSource -match '\$global:OutputEncoding\s*=') 'PowerShell 主程序必须在输出中文前启用 UTF-8。'
Assert-True ($mainScriptSource -match 'Get-ChromiumExtensions\s+-IncludeInactive:\$IncludeInactiveValue\s+-AllUsers:\$AllUsersValue') 'Chromium 插件扫描必须正确接收“扫描所有用户”选项。'
Assert-True ($mainScriptSource -match 'Get-FirefoxExtensions\s+-IncludeInactive:\$IncludeInactiveValue\s+-AllUsers:\$AllUsersValue') 'Firefox 插件扫描必须正确接收“扫描所有用户”选项。'
$extensionSource = Get-Content -LiteralPath (Join-Path $root 'Get-InstalledExtensions.ps1') -Raw -Encoding UTF8
foreach ($supportedBrowserPath in @('Google\Chrome\User Data', 'Microsoft\Edge\User Data', 'Chromium\User Data', 'BraveSoftware\Brave-Browser\User Data', 'Vivaldi\User Data', 'Naver\Naver Whale\User Data')) {
    Assert-True ($extensionSource -match [regex]::Escape($supportedBrowserPath)) "插件扫描缺少指定浏览器路径：$supportedBrowserPath"
}
Assert-True ($extensionSource -notmatch 'Opera Software|YandexBrowser|Chrome Beta|Edge Beta') '插件扫描不得扩展到用户未指定的其他浏览器。'
Assert-True (Test-Path -LiteralPath (Join-Path $root 'Modules\ImportExcel\7.8.10\ImportExcel.psd1')) '发布前必须包含离线 ImportExcel 7.8.10。'

$node = Get-Command node -ErrorAction SilentlyContinue
if ($null -ne $node) {
    & $node.Source (Join-Path $PSScriptRoot 'check-html.js')
    Assert-True ($LASTEXITCODE -eq 0) 'HTML 页面 JavaScript 语法检查失败。'
}

if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Host ('[失败] ' + $_) -ForegroundColor Red }
    throw "$($failures.Count) 项测试失败。"
}
Write-Host 'CheckSentry 自动化检查全部通过。' -ForegroundColor Green
