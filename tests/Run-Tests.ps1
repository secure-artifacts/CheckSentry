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
        $fixturePackage.Save()
    } finally { $fixturePackage.Dispose() }

    $cloudValidationError = ''
    try { $null = Import-CloudRuleSheets -Path $cloudFixturePath } catch { $cloudValidationError = [string]$_.Exception.Message }
    Assert-True ($cloudValidationError -match '白名单.*F2.*版本号.*2049') '云端规则校验错误必须指出具体工作表、单元格、字段和值长度。'
    Assert-True ($cloudValidationError -notmatch '其他数据') '云端导入必须忽略待定、白名单、黑名单以外的工作表。'
} finally {
    if (Test-Path -LiteralPath $cloudFixturePath -PathType Leaf) { Remove-Item -LiteralPath $cloudFixturePath -Force }
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
