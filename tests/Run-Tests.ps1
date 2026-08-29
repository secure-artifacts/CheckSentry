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
Assert-True ((Get-SafeChildPath -BasePath $root -RelativePath '..\outside.txt') -eq '') '安全路径函数必须拒绝目录穿越。'

$videoKitDefaults = Get-DefaultRuleFields -Item ([PSCustomObject]@{ 类型='软件'; 名称='VideoKit 4.4.27'; 版本='4.4.27' })
Assert-True ($videoKitDefaults.MatchType -eq '包含') '名称末尾包含版本的软件应默认使用包含匹配。'
Assert-True ($videoKitDefaults.NamePattern -eq 'VideoKit') '软件规则关键词应自动移除名称末尾的版本号。'
Assert-True ($videoKitDefaults.Version -eq '4.4.27') '复制或批量添加规则时必须自动填写当前版本号。'

$legacyVideoKitRule = [PSCustomObject]@{ 类型='软件'; 匹配方式='精确'; 软件名关键词='VideoKit 4.4.28'; 插件ID=''; 版本号=''; 发布者='' }
$legacyVideoKitMatcher = New-RuleMatcher -Rules @($legacyVideoKitRule) -ItemType '软件'
Assert-True ($null -eq (Find-FirstMatchingRule -Rules $legacyVideoKitMatcher -ItemType '软件' -DisplayName 'VideoKit 4.4.27' -Publisher '' -ExtensionId '' -Version '4.4.27')) '旧规则中内嵌的目标版本不应误判为当前版本匹配。'
Assert-True ($null -ne (Find-FirstMatchingRule -Rules $legacyVideoKitMatcher -ItemType '软件' -DisplayName 'VideoKit 4.4.27' -Publisher '' -ExtensionId '' -Version '4.4.27' -IgnoreVersion)) '旧规则名称中的版本号应被自动识别并触发版本变化。'

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
