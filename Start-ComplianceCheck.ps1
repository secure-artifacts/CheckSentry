<#
.SYNOPSIS
    CheckSentry 本地 Windows IT 合规核对工具。
    黑名单优先级最高，白名单次之，待定最低。所有清单写入均通过临时副本校验后提交。
#>

param(
    [string]$ListPath = '',
    [int]$Port = 8787,
    [switch]$IncludeSystemComponents
)

$ErrorActionPreference = 'Stop'
$script:AllowedSheets = @('白名单', '待定', '黑名单')
$script:AllowedItemTypes = @('软件', 'Chromium插件', 'Firefox插件')
$script:AllowedMatchTypes = @('插件ID精确', '精确', '包含', '通配符', '正则')
$script:LastUndoSnapshot = $null
$script:InventoryCache = $null
$script:MaxRequestChars = 1048576

function Initialize-ImportExcel {
    $requiredVersion = [version]'7.8.10'
    $available = Get-Module -ListAvailable -Name ImportExcel |
        Where-Object { $_.Version -eq $requiredVersion } |
        Select-Object -First 1
    if ($null -eq $available) {
        Write-Host "未找到已验证的 ImportExcel $requiredVersion，正在安装到当前用户..." -ForegroundColor Yellow
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
            $nuget = Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($null -eq $nuget) {
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force -ErrorAction Stop | Out-Null
            }
            Install-Module ImportExcel -RequiredVersion $requiredVersion -Repository PSGallery -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        } catch {
            throw "无法安装 ImportExcel $requiredVersion。请联网后手动执行 Install-Module ImportExcel -RequiredVersion $requiredVersion -Scope CurrentUser。详细错误：$($_.Exception.Message)"
        }
    }
    try {
        Import-Module ImportExcel -RequiredVersion $requiredVersion -Force -ErrorAction Stop
        foreach ($commandName in @('Open-ExcelPackage', 'Close-ExcelPackage')) {
            if ($null -eq (Get-Command $commandName -ErrorAction SilentlyContinue)) {
                throw "ImportExcel 缺少命令 $commandName"
            }
        }
    } catch {
        throw "ImportExcel $requiredVersion 加载失败：$($_.Exception.Message)"
    }
}

function Get-ExpectedHeaders {
    param([string]$SheetName)
    if ($SheetName -eq '黑名单') {
        return @('id', '类型', '插件ID', '匹配方式', '软件名关键词', '版本号', '发布者', '分类', '禁止原因', '添加人', '添加时间')
    }
    return @('id', '类型', '插件ID', '匹配方式', '软件名关键词', '版本号', '发布者', '状态', '分类', '备注', '添加人', '添加时间')
}

function Assert-AllowedSheet {
    param([object]$SheetName)
    $name = [string]$SheetName
    if ($script:AllowedSheets -notcontains $name) { throw "不允许的工作表：$name" }
    return $name
}

function Assert-WorkbookPackageSchema {
    param($Package)
    $expectedSheetNames = @('说明') + $script:AllowedSheets
    $actualSheetNames = @($Package.Workbook.Worksheets | ForEach-Object { $_.Name })
    if ($actualSheetNames.Count -ne $expectedSheetNames.Count) {
        throw "清单必须且只能包含 4 个工作表：说明、白名单、待定、黑名单。"
    }
    foreach ($actualSheetName in $actualSheetNames) {
        if ($expectedSheetNames -notcontains $actualSheetName) { throw "清单包含未定义的工作表【$actualSheetName】。" }
    }
    foreach ($sheetName in $script:AllowedSheets) {
        $worksheet = $Package.Workbook.Worksheets[$sheetName]
        if ($null -eq $worksheet) { throw "清单缺少工作表【$sheetName】。" }
        $headers = Get-ExpectedHeaders -SheetName $sheetName
        if ($null -ne $worksheet.Dimension -and $worksheet.Dimension.End.Row -gt 100000) {
            throw "工作表【$sheetName】超过 100000 行，已拒绝处理。"
        }
        if ($null -ne $worksheet.Dimension -and (([long]$worksheet.Dimension.End.Row * [long]$worksheet.Dimension.End.Column) -gt 2000000)) {
            throw "工作表【$sheetName】使用范围异常大，已拒绝处理。"
        }
        for ($column = 1; $column -le $headers.Count; $column++) {
            $actual = [string]($worksheet.Cells[1, $column].Text)
            if ($actual -ne $headers[$column - 1]) {
                throw "工作表【$sheetName】第 $column 列应为【$($headers[$column - 1])】，实际为【$actual】。"
            }
        }
        if ($null -ne $worksheet.Dimension) {
            $seenIds = @{}
            for ($row = 1; $row -le $worksheet.Dimension.End.Row; $row++) {
                for ($column = 1; $column -le $worksheet.Dimension.End.Column; $column++) {
                    $cell = $worksheet.Cells[$row, $column]
                    if (-not [string]::IsNullOrWhiteSpace([string]$cell.Formula)) {
                        throw "工作表【$sheetName】单元格 $($cell.Address) 含公式。规则清单只允许纯文本/数值。"
                    }
                    if ($column -gt $headers.Count -and -not [string]::IsNullOrWhiteSpace([string]$cell.Value)) {
                        throw "工作表【$sheetName】存在未定义的第 $column 列数据，已拒绝处理。"
                    }
                }
                if ($row -ge 2) {
                    $rowHasValue = $false
                    for ($column = 1; $column -le $headers.Count; $column++) {
                        if (-not [string]::IsNullOrWhiteSpace([string]($worksheet.Cells[$row, $column].Value))) { $rowHasValue = $true; break }
                    }
                    if ($rowHasValue) {
                        $id = [string]($worksheet.Cells[$row, 1].Value)
                        if ($id -notmatch '^\d{1,10}$') { throw "工作表【$sheetName】第 $row 行规则 ID 无效。" }
                        if ($seenIds.ContainsKey($id)) { throw "工作表【$sheetName】存在重复规则 ID：$id。" }
                        $seenIds[$id] = $true
                    }
                }
            }
        }
    }
    $explanation = $Package.Workbook.Worksheets['说明']
    if ($null -eq $explanation) { throw '清单缺少工作表【说明】。' }
    if ($null -ne $explanation.Dimension) {
        if (($explanation.Dimension.End.Row * $explanation.Dimension.End.Column) -gt 100000) { throw '工作表【说明】内容范围异常大，已拒绝处理。' }
        for ($row = 1; $row -le $explanation.Dimension.End.Row; $row++) {
            for ($column = 1; $column -le $explanation.Dimension.End.Column; $column++) {
                $cell = $explanation.Cells[$row, $column]
                if (-not [string]::IsNullOrWhiteSpace([string]$cell.Formula)) {
                    throw "工作表【说明】单元格 $($cell.Address) 含公式。说明工作表只允许纯文本。"
                }
            }
        }
    }
}

function Assert-WorkbookPathSchema {
    param([string]$Path)
    $package = $null
    try {
        $package = Open-ExcelPackage -Path $Path -ErrorAction Stop
        Assert-WorkbookPackageSchema -Package $package
    } finally {
        if ($null -ne $package) { $package.Dispose() }
    }
}

function Convert-WorksheetToRules {
    param($Worksheet, [string]$SheetName)
    $headers = Get-ExpectedHeaders -SheetName $SheetName
    $rules = @()
    if ($null -eq $Worksheet.Dimension -or $Worksheet.Dimension.End.Row -lt 2) { return $rules }
    for ($row = 2; $row -le $Worksheet.Dimension.End.Row; $row++) {
        $hasValue = $false
        $ordered = [ordered]@{}
        for ($column = 1; $column -le $headers.Count; $column++) {
            $cell = $Worksheet.Cells[$row, $column]
            if (-not [string]::IsNullOrWhiteSpace([string]$cell.Formula)) {
                throw "工作表【$SheetName】单元格 $($cell.Address) 含公式。规则清单只允许纯文本/数值。"
            }
            $value = $cell.Value
            if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) { $hasValue = $true }
            $ordered[$headers[$column - 1]] = $value
        }
        if ($hasValue) { $rules += [PSCustomObject]$ordered }
    }
    return $rules
}

function Import-RuleSheet {
    param([string]$Path, [string]$SheetName)
    $null = Assert-AllowedSheet -SheetName $SheetName
    $lastError = $null
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        $package = $null
        try {
            $package = Open-ExcelPackage -Path $Path -ErrorAction Stop
            Assert-WorkbookPackageSchema -Package $package
            return @(Convert-WorksheetToRules -Worksheet $package.Workbook.Worksheets[$SheetName] -SheetName $SheetName)
        } catch {
            $lastError = $_
            if ($attempt -lt 3) { Start-Sleep -Milliseconds 350 }
        } finally {
            if ($null -ne $package) { $package.Dispose() }
        }
    }
    throw "读取【$SheetName】失败。请确认 Excel 已关闭且文件未损坏。详细错误：$($lastError.Exception.Message)"
}

function Test-ListFileAvailable {
    param([string]$Path)
    $stream = $null
    try {
        $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
    } catch {
        throw "清单文件正被其他程序占用或不可写：$Path"
    } finally {
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

function Get-FileFingerprint {
    param([string]$Path)
    $item = Get-Item -LiteralPath $Path -ErrorAction Stop
    $hash = Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop
    return ('{0}|{1}|{2}' -f $item.Length, $item.LastWriteTimeUtc.Ticks, $hash.Hash)
}

function Remove-UndoSnapshot {
    if ($null -ne $script:LastUndoSnapshot) {
        try {
            if (Test-Path -LiteralPath $script:LastUndoSnapshot.SnapshotPath -PathType Leaf) {
                Remove-Item -LiteralPath $script:LastUndoSnapshot.SnapshotPath -Force -ErrorAction SilentlyContinue
            }
        } catch {
            Write-Verbose "清理撤销快照失败：$($_.Exception.Message)"
        }
    }
    $script:LastUndoSnapshot = $null
}

function New-UndoSnapshotFile {
    param([string]$Path)
    $undoDirectory = Join-Path ([System.IO.Path]::GetTempPath()) 'CheckSentry-Undo'
    if (-not (Test-Path -LiteralPath $undoDirectory -PathType Container)) {
        New-Item -ItemType Directory -Path $undoDirectory -Force | Out-Null
    }
    $snapshotPath = Join-Path $undoDirectory (([guid]::NewGuid().ToString('N')) + '.xlsx')
    Copy-Item -LiteralPath $Path -Destination $snapshotPath -ErrorAction Stop
    return $snapshotPath
}

function Set-WorkbookFileFromTemp {
    param([string]$NewPath, [string]$DestinationPath)
    Test-ListFileAvailable -Path $DestinationPath
    $backupPath = Join-Path ([System.IO.Path]::GetDirectoryName($DestinationPath)) ('.' + [System.IO.Path]::GetFileName($DestinationPath) + '.' + [guid]::NewGuid().ToString('N') + '.backup.xlsx')
    try {
        [System.IO.File]::Replace($NewPath, $DestinationPath, $backupPath, $true)
    } catch {
        try {
            if (-not (Test-Path -LiteralPath $backupPath -PathType Leaf)) {
                Copy-Item -LiteralPath $DestinationPath -Destination $backupPath -ErrorAction Stop
            }
            Move-Item -LiteralPath $NewPath -Destination $DestinationPath -Force -ErrorAction Stop
        } catch {
            if (Test-Path -LiteralPath $backupPath -PathType Leaf) {
                Copy-Item -LiteralPath $backupPath -Destination $DestinationPath -Force -ErrorAction SilentlyContinue
            }
            throw
        }
    } finally {
        if (Test-Path -LiteralPath $backupPath -PathType Leaf) {
            Remove-Item -LiteralPath $backupPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Invoke-WorkbookTransaction {
    param(
        [string]$Path,
        [scriptblock]$Operation,
        [switch]$CreateUndo
    )
    Test-ListFileAvailable -Path $Path
    $fingerprintBefore = Get-FileFingerprint -Path $Path
    $pendingSnapshot = $null
    if ($CreateUndo) { $pendingSnapshot = New-UndoSnapshotFile -Path $Path }

    $directory = [System.IO.Path]::GetDirectoryName($Path)
    $temporaryPath = Join-Path $directory ('.' + [System.IO.Path]::GetFileName($Path) + '.' + [guid]::NewGuid().ToString('N') + '.tmp.xlsx')
    $package = $null
    try {
        Copy-Item -LiteralPath $Path -Destination $temporaryPath -ErrorAction Stop
        $package = Open-ExcelPackage -Path $temporaryPath -ErrorAction Stop
        Assert-WorkbookPackageSchema -Package $package
        $result = & $Operation $package
        Assert-WorkbookPackageSchema -Package $package
        Close-ExcelPackage -ExcelPackage $package -ErrorAction Stop
        $package = $null
        Assert-WorkbookPathSchema -Path $temporaryPath
        if ((Get-FileFingerprint -Path $Path) -ne $fingerprintBefore) {
            throw '清单在操作期间被外部修改，已取消本次写入。请关闭 Excel 后重试。'
        }
        Set-WorkbookFileFromTemp -NewPath $temporaryPath -DestinationPath $Path
        if ($CreateUndo) {
            Remove-UndoSnapshot
            $script:LastUndoSnapshot = [PSCustomObject]@{
                ListPath = $Path
                SnapshotPath = $pendingSnapshot
                ExpectedFingerprint = Get-FileFingerprint -Path $Path
                CreatedAt = Get-Date
            }
            $pendingSnapshot = $null
        }
        return $result
    } finally {
        if ($null -ne $package) { $package.Dispose() }
        if (Test-Path -LiteralPath $temporaryPath -PathType Leaf) { Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue }
        if ($null -ne $pendingSnapshot -and (Test-Path -LiteralPath $pendingSnapshot -PathType Leaf)) {
            Remove-Item -LiteralPath $pendingSnapshot -Force -ErrorAction SilentlyContinue
        }
    }
}

function Restore-UndoSnapshot {
    param([string]$Path)
    if ($null -eq $script:LastUndoSnapshot) { throw '没有可撤销的维护操作。' }
    if ([string]$script:LastUndoSnapshot.ListPath -ne $Path) { throw '撤销快照与当前清单不一致。' }
    if (-not (Test-Path -LiteralPath $script:LastUndoSnapshot.SnapshotPath -PathType Leaf)) { throw '撤销快照已丢失。' }
    if ((Get-FileFingerprint -Path $Path) -ne $script:LastUndoSnapshot.ExpectedFingerprint) {
        throw '检测到清单在上一步操作后被外部修改。为避免覆盖这些修改，已拒绝撤销。'
    }
    $directory = [System.IO.Path]::GetDirectoryName($Path)
    $temporaryPath = Join-Path $directory ('.' + [System.IO.Path]::GetFileName($Path) + '.' + [guid]::NewGuid().ToString('N') + '.undo.xlsx')
    try {
        Copy-Item -LiteralPath $script:LastUndoSnapshot.SnapshotPath -Destination $temporaryPath -ErrorAction Stop
        Assert-WorkbookPathSchema -Path $temporaryPath
        Set-WorkbookFileFromTemp -NewPath $temporaryPath -DestinationPath $Path
        Remove-UndoSnapshot
    } finally {
        if (Test-Path -LiteralPath $temporaryPath -PathType Leaf) { Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue }
    }
}

function Get-LimitedText {
    param([object]$Value, [string]$FieldName, [int]$MaximumLength, [switch]$Required)
    $text = if ($null -eq $Value) { '' } else { [string]$Value }
    $text = $text -replace "`0", ''
    if ($Required -and [string]::IsNullOrWhiteSpace($text)) { throw "$FieldName 不能为空。" }
    if ($text.Length -gt $MaximumLength) { throw "$FieldName 不能超过 $MaximumLength 个字符。" }
    return $text
}

function Convert-InputToRule {
    param($Data, [string]$TargetSheet)
    $sheet = Assert-AllowedSheet -SheetName $TargetSheet
    $itemType = Get-LimitedText -Value $Data.type -FieldName '类型' -MaximumLength 32 -Required
    if ($script:AllowedItemTypes -notcontains $itemType) { throw "不允许的类型：$itemType" }
    $matchType = Get-LimitedText -Value $Data.matchType -FieldName '匹配方式' -MaximumLength 32 -Required
    if ($script:AllowedMatchTypes -notcontains $matchType) { throw "不允许的匹配方式：$matchType" }
    $extensionId = Get-LimitedText -Value $Data.extId -FieldName '插件ID' -MaximumLength 256
    $namePattern = Get-LimitedText -Value $Data.namePattern -FieldName '软件名关键词' -MaximumLength 512 -Required
    if ($itemType -eq '软件' -and $matchType -eq '插件ID精确') { throw '软件规则不能使用插件ID精确。' }
    if ($matchType -eq '插件ID精确' -and [string]::IsNullOrWhiteSpace($extensionId)) { throw '插件ID精确规则必须填写插件ID。' }
    if ($matchType -eq '正则') {
        try {
            $regex = New-Object System.Text.RegularExpressions.Regex($namePattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase, [TimeSpan]::FromMilliseconds(250))
            $null = $regex.IsMatch('CheckSentry validation')
        } catch {
            throw "正则表达式无效或执行超时：$($_.Exception.Message)"
        }
    }
    $rule = [ordered]@{
        类型 = $itemType
        插件ID = $extensionId
        匹配方式 = $matchType
        软件名关键词 = $namePattern
        版本号 = Get-LimitedText -Value $Data.version -FieldName '版本号' -MaximumLength 128
        发布者 = Get-LimitedText -Value $Data.publisher -FieldName '发布者' -MaximumLength 256
        分类 = Get-LimitedText -Value $Data.category -FieldName '分类' -MaximumLength 128
        添加人 = Get-LimitedText -Value $env:USERNAME -FieldName '添加人' -MaximumLength 128
        添加时间 = Get-Date -Format 'yyyy-MM-dd HH:mm'
    }
    $note = Get-LimitedText -Value $Data.note -FieldName '备注' -MaximumLength 2000
    if ($sheet -eq '黑名单') { $rule.禁止原因 = $note }
    elseif ($sheet -eq '待定') { $rule.状态 = '待定'; $rule.备注 = $note }
    else { $rule.状态 = '允许'; $rule.备注 = $note }
    return [PSCustomObject]$rule
}

function Get-RuleIdentity {
    param([string]$ItemType, [string]$ExtensionId, [string]$NamePattern)
    if ($ItemType -eq '软件') { return ('软件|' + ([string]$NamePattern).Trim().ToLowerInvariant()) }
    return ($ItemType + '|' + ([string]$ExtensionId).Trim().ToLowerInvariant())
}

function Get-WorksheetRowIdentity {
    param($Worksheet, [int]$Row, [string]$SheetName)
    $headers = Get-ExpectedHeaders -SheetName $SheetName
    $map = @{}
    for ($column = 1; $column -le $headers.Count; $column++) { $map[$headers[$column - 1]] = [string]($Worksheet.Cells[$Row, $column].Value) }
    $itemType = if ([string]::IsNullOrWhiteSpace($map['类型'])) { '软件' } else { $map['类型'] }
    return Get-RuleIdentity -ItemType $itemType -ExtensionId $map['插件ID'] -NamePattern $map['软件名关键词']
}

function Remove-RuleByIdFromPackage {
    param($Package, [string]$SheetName, [string]$RuleId)
    $sheet = Assert-AllowedSheet -SheetName $SheetName
    if ($RuleId -notmatch '^\d{1,10}$') { throw '规则 ID 无效。' }
    $worksheet = $Package.Workbook.Worksheets[$sheet]
    if ($null -eq $worksheet.Dimension) { throw "工作表【$sheet】中不存在规则 $RuleId。" }
    for ($row = 2; $row -le $worksheet.Dimension.End.Row; $row++) {
        if ([string]($worksheet.Cells[$row, 1].Value) -eq $RuleId) {
            $worksheet.DeleteRow($row, 1)
            return
        }
    }
    throw "工作表【$sheet】中不存在规则 $RuleId。"
}

function Remove-IdentityFromPackage {
    param($Package, [string]$Identity)
    foreach ($sheetName in $script:AllowedSheets) {
        $worksheet = $Package.Workbook.Worksheets[$sheetName]
        if ($null -eq $worksheet.Dimension) { continue }
        for ($row = $worksheet.Dimension.End.Row; $row -ge 2; $row--) {
            if ((Get-WorksheetRowIdentity -Worksheet $worksheet -Row $row -SheetName $sheetName) -eq $Identity) {
                $worksheet.DeleteRow($row, 1)
            }
        }
    }
}

function Test-IdentityInSheet {
    param($Package, [string]$SheetName, [string]$Identity)
    $worksheet = $Package.Workbook.Worksheets[$SheetName]
    if ($null -eq $worksheet.Dimension) { return $false }
    for ($row = 2; $row -le $worksheet.Dimension.End.Row; $row++) {
        if ((Get-WorksheetRowIdentity -Worksheet $worksheet -Row $row -SheetName $SheetName) -eq $Identity) { return $true }
    }
    return $false
}

function Add-RuleToPackage {
    param($Package, [string]$SheetName, $Rule)
    $sheet = Assert-AllowedSheet -SheetName $SheetName
    $worksheet = $Package.Workbook.Worksheets[$sheet]
    $headers = Get-ExpectedHeaders -SheetName $sheet
    $maximumId = [long]0
    $targetRow = 2
    if ($null -ne $worksheet.Dimension) {
        for ($row = 2; $row -le $worksheet.Dimension.End.Row; $row++) {
            $parsedId = [long]0
            if ([long]::TryParse([string]($worksheet.Cells[$row, 1].Value), [ref]$parsedId) -and $parsedId -gt $maximumId) { $maximumId = $parsedId }
            $rowHasValue = $false
            for ($column = 1; $column -le $headers.Count; $column++) {
                if (-not [string]::IsNullOrWhiteSpace([string]($worksheet.Cells[$row, $column].Value))) { $rowHasValue = $true; break }
            }
            if ($rowHasValue) { $targetRow = $row + 1 }
        }
    }
    if ($maximumId -ge 9999999999) { throw "工作表【$sheet】的规则 ID 已达到上限。" }
    for ($column = 1; $column -le $headers.Count; $column++) {
        $header = $headers[$column - 1]
        if ($header -eq 'id') { $worksheet.Cells[$targetRow, $column].Value = [long]($maximumId + 1) }
        else {
            $property = $Rule.PSObject.Properties[$header]
            $value = if ($null -ne $property -and $null -ne $property.Value) { [string]$property.Value } else { '' }
            $worksheet.Cells[$targetRow, $column].Value = $value
            $worksheet.Cells[$targetRow, $column].Style.Numberformat.Format = '@'
        }
    }
    return ($maximumId + 1)
}

function Test-VersionMatch {
    param([string]$InstalledVersion, [string]$RuleVersion)
    if ([string]::IsNullOrWhiteSpace($RuleVersion)) { return $true }
    if ([string]::IsNullOrWhiteSpace($InstalledVersion)) { return $false }
    return $InstalledVersion -like $RuleVersion
}

function Test-NameMatch {
    param([string]$InstalledName, $Rule)
    if ($Rule.'匹配方式' -eq '插件ID精确') {
        if ([string]::IsNullOrWhiteSpace([string]$Rule.'插件ID')) { return $false }
        return [string]::Equals($InstalledName, [string]$Rule.'插件ID', [System.StringComparison]::OrdinalIgnoreCase)
    }
    $pattern = [string]$Rule.'软件名关键词'
    if ([string]::IsNullOrWhiteSpace($pattern)) { return $false }
    switch ([string]$Rule.'匹配方式') {
        '精确' { return [string]::Equals($InstalledName, $pattern, [System.StringComparison]::OrdinalIgnoreCase) }
        '包含' { return $InstalledName.IndexOf($pattern, [System.StringComparison]::OrdinalIgnoreCase) -ge 0 }
        '通配符' { return $InstalledName -like $pattern }
        '正则' {
            try { return [regex]::IsMatch($InstalledName, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase, [TimeSpan]::FromMilliseconds(250)) }
            catch { return $false }
        }
        default { return $false }
    }
}

function Test-PublisherMatch {
    param([string]$InstalledPublisher, [object]$RulePublisher)
    $publisher = [string]$RulePublisher
    if ([string]::IsNullOrWhiteSpace($publisher)) { return $true }
    if ([string]::IsNullOrWhiteSpace($InstalledPublisher)) { return $false }
    return $InstalledPublisher.IndexOf($publisher, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
}

function Find-FirstMatchingRule {
    param($Rules, [string]$ItemType, [string]$DisplayName, [string]$Publisher, [string]$ExtensionId)
    foreach ($rule in @($Rules)) {
        $ruleType = if ([string]::IsNullOrWhiteSpace([string]$rule.'类型')) { '软件' } else { [string]$rule.'类型' }
        if ($ruleType -ne $ItemType) { continue }
        $testName = if ($rule.'匹配方式' -eq '插件ID精确') { $ExtensionId } else { $DisplayName }
        if (-not (Test-NameMatch -InstalledName $testName -Rule $rule)) { continue }
        if (-not (Test-PublisherMatch -InstalledPublisher $Publisher -RulePublisher $rule.'发布者')) { continue }
        return $rule
    }
    return $null
}

function New-ComplianceItem {
    param($Source, [string]$ItemType, [string]$Status, [string]$Reason, $MatchedRule, [string]$MatchedSheet)
    $isSoftware = $ItemType -eq '软件'
    return [PSCustomObject]@{
        名称 = if ($isSoftware) { [string]$Source.名称 } else { [string]$Source.Name }
        版本 = if ($isSoftware) { [string]$Source.版本 } else { [string]$Source.Version }
        发布者 = if ($isSoftware) { [string]$Source.发布者 } else { [string]$Source.Publisher }
        安装日期 = if ($isSoftware) { [string]$Source.安装日期 } else { '' }
        安装路径 = if ($isSoftware) { [string]$Source.安装路径 } else { '' }
        状态 = $Status
        原因 = $Reason
        类型 = $ItemType
        插件ID = if ($isSoftware) { '' } else { [string]$Source.ExtensionId }
        Locations = if ($isSoftware) { $null } else { $Source.Locations }
        图标路径 = if ($isSoftware) { [string]$Source.图标路径 } else { [string]$Source.IconPath }
        MatchedSheet = $MatchedSheet
        MatchedRuleId = if ($null -ne $MatchedRule) { $MatchedRule.id } else { $null }
    }
}

function Get-ComplianceResult {
    param($Installed, $BlackRules, $WhiteRules, $PendingRules, [string]$ItemType)
    $results = @()
    foreach ($item in @($Installed)) {
        $displayName = if ($ItemType -eq '软件') { [string]$item.名称 } else { [string]$item.Name }
        $publisher = if ($ItemType -eq '软件') { [string]$item.发布者 } else { [string]$item.Publisher }
        $version = if ($ItemType -eq '软件') { [string]$item.版本 } else { [string]$item.Version }
        $extensionId = if ($ItemType -eq '软件') { '' } else { [string]$item.ExtensionId }

        $black = Find-FirstMatchingRule -Rules $BlackRules -ItemType $ItemType -DisplayName $displayName -Publisher $publisher -ExtensionId $extensionId
        if ($null -ne $black) {
            $reason = if ([string]::IsNullOrWhiteSpace([string]$black.'禁止原因')) { '命中黑名单' } else { [string]$black.'禁止原因' }
            $results += New-ComplianceItem -Source $item -ItemType $ItemType -Status '命中黑名单' -Reason $reason -MatchedRule $black -MatchedSheet '黑名单'
            continue
        }

        $white = Find-FirstMatchingRule -Rules $WhiteRules -ItemType $ItemType -DisplayName $displayName -Publisher $publisher -ExtensionId $extensionId
        if ($null -ne $white) {
            if (Test-VersionMatch -InstalledVersion $version -RuleVersion ([string]$white.'版本号')) {
                $reason = if ([string]::IsNullOrWhiteSpace([string]$white.'备注')) { '' } else { [string]$white.'备注' }
                $results += New-ComplianceItem -Source $item -ItemType $ItemType -Status '已匹配' -Reason $reason -MatchedRule $white -MatchedSheet '白名单'
            } else {
                $reason = "清单登记版本为 [$($white.'版本号')]，当前安装版本为 [$version]"
                $results += New-ComplianceItem -Source $item -ItemType $ItemType -Status '版本变化' -Reason $reason -MatchedRule $white -MatchedSheet '白名单'
            }
            continue
        }

        $pending = Find-FirstMatchingRule -Rules $PendingRules -ItemType $ItemType -DisplayName $displayName -Publisher $publisher -ExtensionId $extensionId
        if ($null -ne $pending) {
            $reason = if ([string]::IsNullOrWhiteSpace([string]$pending.'备注')) { '待定' } else { [string]$pending.'备注' }
            $results += New-ComplianceItem -Source $item -ItemType $ItemType -Status '待定' -Reason $reason -MatchedRule $pending -MatchedSheet '待定'
            continue
        }

        $results += New-ComplianceItem -Source $item -ItemType $ItemType -Status '待定' -Reason '扫描到的新项目，待归类' -MatchedRule $null -MatchedSheet ''
    }
    return $results
}

function Add-NewPendingRules {
    param([string]$Path, [object[]]$Results)
    $candidates = @($Results | Where-Object { $_.状态 -eq '待定' -and $null -eq $_.MatchedRuleId })
    if ($candidates.Count -eq 0) { return 0 }
    $operation = {
        param($package)
        $added = 0
        foreach ($item in $candidates) {
            $itemType = [string]$item.类型
            $identity = Get-RuleIdentity -ItemType $itemType -ExtensionId ([string]$item.插件ID) -NamePattern ([string]$item.名称)
            $exists = $false
            foreach ($sheetName in $script:AllowedSheets) {
                if (Test-IdentityInSheet -Package $package -SheetName $sheetName -Identity $identity) { $exists = $true; break }
            }
            if ($exists) { continue }
            $ruleData = [PSCustomObject]@{
                type = $itemType
                extId = [string]$item.插件ID
                matchType = if ($itemType -eq '软件') { '精确' } else { '插件ID精确' }
                namePattern = [string]$item.名称
                version = ''
                publisher = [string]$item.发布者
                category = ''
                note = '扫描自动加入待定'
            }
            $rule = Convert-InputToRule -Data $ruleData -TargetSheet '待定'
            $null = Add-RuleToPackage -Package $package -SheetName '待定' -Rule $rule
            $added++
        }
        return $added
    }
    $count = Invoke-WorkbookTransaction -Path $Path -Operation $operation
    if ($count -gt 0) { Remove-UndoSnapshot }
    return $count
}

function Get-InventoryData {
    param([switch]$ForceRefresh, [switch]$IncludeSystem)
    if (-not $ForceRefresh -and $null -ne $script:InventoryCache -and ((Get-Date) - $script:InventoryCache.CreatedAt).TotalSeconds -lt 30) {
        return $script:InventoryCache
    }
    $script:InventoryCache = [PSCustomObject]@{
        Software = @(Get-InstalledSoftwareList -IncludeSystemComponents:$IncludeSystem)
        Chromium = @(Get-ChromiumExtensions)
        Firefox = @(Get-FirefoxExtensions)
        CreatedAt = Get-Date
    }
    return $script:InventoryCache
}

function Get-LocalIconDataUri {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) { return '' }
    try {
        $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
        if ($extension -in @('.exe', '.dll')) {
            Add-Type -AssemblyName System.Drawing -ErrorAction Stop
            $icon = $null; $bitmap = $null; $stream = $null
            try {
                $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($Path)
                if ($null -eq $icon) { return '' }
                $bitmap = $icon.ToBitmap()
                $stream = New-Object System.IO.MemoryStream
                $bitmap.Save($stream, [System.Drawing.Imaging.ImageFormat]::Png)
                return 'data:image/png;base64,' + [Convert]::ToBase64String($stream.ToArray())
            } finally {
                if ($null -ne $stream) { $stream.Dispose() }
                if ($null -ne $bitmap) { $bitmap.Dispose() }
                if ($null -ne $icon) { $icon.Dispose() }
            }
        }
        if ($extension -notin @('.png', '.jpg', '.jpeg', '.gif', '.ico')) { return '' }
        $bytes = [System.IO.File]::ReadAllBytes($Path)
        if ($bytes.Length -eq 0 -or $bytes.Length -gt 524288) { return '' }
        $mime = ''
        if ($extension -eq '.png' -and $bytes.Length -ge 8 -and $bytes[0] -eq 0x89 -and $bytes[1] -eq 0x50 -and $bytes[2] -eq 0x4E -and $bytes[3] -eq 0x47) { $mime = 'image/png' }
        elseif ($extension -in @('.jpg', '.jpeg') -and $bytes.Length -ge 3 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xD8 -and $bytes[2] -eq 0xFF) { $mime = 'image/jpeg' }
        elseif ($extension -eq '.gif' -and $bytes.Length -ge 6 -and [Text.Encoding]::ASCII.GetString($bytes, 0, 6) -match '^GIF8[79]a$') { $mime = 'image/gif' }
        elseif ($extension -eq '.ico' -and $bytes.Length -ge 4 -and $bytes[0] -eq 0 -and $bytes[1] -eq 0 -and $bytes[2] -eq 1 -and $bytes[3] -eq 0) { $mime = 'image/x-icon' }
        if ([string]::IsNullOrWhiteSpace($mime)) { return '' }
        return 'data:' + $mime + ';base64,' + [Convert]::ToBase64String($bytes)
    } catch { return '' }
}

function Get-RuleIconDataUri {
    param($Rule, [bool]$IncludeSystem)
    $inventory = Get-InventoryData -IncludeSystem:$IncludeSystem
    $itemType = if ([string]::IsNullOrWhiteSpace([string]$Rule.'类型')) { '软件' } else { [string]$Rule.'类型' }
    $iconPath = ''
    if ($itemType -eq '软件') {
        $software = $inventory.Software | Where-Object {
            Test-NameMatch -InstalledName ([string]$_.名称) -Rule $Rule
        } | Select-Object -First 1
        if ($null -ne $software) { $iconPath = [string]$software.图标路径 }
    } else {
        $extensions = if ($itemType -eq 'Firefox插件') { $inventory.Firefox } else { $inventory.Chromium }
        $extension = $extensions | Where-Object {
            [string]::Equals([string]$_.ExtensionId, [string]$Rule.'插件ID', [System.StringComparison]::OrdinalIgnoreCase)
        } | Select-Object -First 1
        if ($null -ne $extension) { $iconPath = [string]$extension.IconPath }
    }
    return Get-LocalIconDataUri -Path $iconPath
}

function Convert-RulesForWeb {
    param([object[]]$Rules, [bool]$IncludeSystem)
    $output = @()
    foreach ($rule in @($Rules)) {
        $ordered = [ordered]@{}
        foreach ($property in $rule.PSObject.Properties) { $ordered[$property.Name] = $property.Value }
        $ordered['图标数据'] = Get-RuleIconDataUri -Rule $rule -IncludeSystem $IncludeSystem
        $output += [PSCustomObject]$ordered
    }
    return $output
}

function ConvertTo-HtmlEncodedText {
    param([object]$Value)
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Get-RowHtml {
    param($Item, [bool]$NeedsAction)
    $name = ConvertTo-HtmlEncodedText $Item.名称
    $version = ConvertTo-HtmlEncodedText $Item.版本
    $publisher = ConvertTo-HtmlEncodedText $Item.发布者
    $reason = ConvertTo-HtmlEncodedText $Item.原因
    $status = ConvertTo-HtmlEncodedText $Item.状态
    $typeIcon = if ($Item.类型 -eq '软件') { '&#128187;' } elseif ($Item.类型 -eq 'Chromium插件') { '&#127760;' } else { '&#129418;' }
    $iconSource = Get-LocalIconDataUri -Path ([string]$Item.图标路径)
    $iconHtml = if ($iconSource) { "<img class='item-icon' src='$(ConvertTo-HtmlEncodedText $iconSource)' alt='' loading='lazy'><span class='type-fallback hidden'>$typeIcon</span>" } else { "<span class='type-fallback'>$typeIcon</span>" }

    $locationText = ''
    if ($Item.类型 -eq '软件') {
        if (-not [string]::IsNullOrWhiteSpace([string]$Item.安装路径)) { $locationText = '安装位置: ' + [string]$Item.安装路径 }
    } elseif ($Item.Locations) {
        $parts = @()
        foreach ($location in @($Item.Locations)) {
            $browser = if ($location.Browser) { [string]$location.Browser } elseif ($Item.类型 -eq 'Firefox插件') { 'Firefox' } else { '' }
            $locationProfile = [string]$location.ProfileName
            if ($browser -and $locationProfile) { $parts += ($browser + ' / ' + $locationProfile) }
            elseif ($browser) { $parts += $browser }
            elseif ($locationProfile) { $parts += $locationProfile }
        }
        $locationText = '安装位置: ' + (($parts | Select-Object -Unique) -join ' · ')
    }
    $locationHtml = if ([string]::IsNullOrWhiteSpace($locationText)) { '' } else { "<div class='item-location'>$(ConvertTo-HtmlEncodedText $locationText)</div>" }
    $rowClass = switch ([string]$Item.状态) {
        '命中黑名单' { 'row-red row-banned' }
        '待定' { 'row-pending' }
        '版本变化' { 'row-yellow' }
        default { 'row-green' }
    }

    $actionHtml = ''
    if ($NeedsAction) {
        $matchOptions = if ($Item.类型 -eq '软件') {
            '<option value="精确" selected>精确</option><option value="包含">包含</option><option value="通配符">通配符</option><option value="正则">正则</option>'
        } else {
            '<option value="插件ID精确" selected>插件ID精确</option><option value="包含">包含(按名称)</option><option value="通配符">通配符(按名称)</option><option value="正则">正则(按名称)</option>'
        }
        $actionHtml = @"
        <details><summary>操作 &#9662;</summary><form class="rule-action-form">
          <input type="hidden" name="extId" value="$(ConvertTo-HtmlEncodedText $Item.插件ID)">
          <input type="hidden" name="itemType" value="$(ConvertTo-HtmlEncodedText $Item.类型)">
          <input type="hidden" name="matchedSheet" value="$(ConvertTo-HtmlEncodedText $Item.MatchedSheet)">
          <input type="hidden" name="matchedRuleId" value="$(ConvertTo-HtmlEncodedText $Item.MatchedRuleId)">
          <label>匹配方式<select name="matchType">$matchOptions</select></label>
          <label>关键词<input type="text" name="namePattern" value="$name" maxlength="512"></label>
          <label>版本号（留空=不锁版本）<input type="text" name="version" value="" maxlength="128"><button type="button" class="use-current-version" data-current-version="$version">用当前版本</button></label>
          <label>发布者<input type="text" name="publisher" value="$publisher" maxlength="256"></label>
          <label>分类<input type="text" name="category" value="" maxlength="128"></label>
          <label>备注<input type="text" name="note" value="" maxlength="2000"></label>
          <div class="btnrow"><button type="submit" data-status="允许" class="btn-approve">加入白名单（允许）</button><button type="submit" data-status="禁止" class="btn-ban">加入黑名单（禁止）</button><button type="submit" data-status="待定" class="btn-pending">标记待定</button></div>
        </form></details>
"@
    }
    $checkbox = if ($NeedsAction) { "<input type='checkbox' class='row-chk'>" } else { '' }
    return @"
    <tr class="$rowClass"><td class="chk-cell">$checkbox</td><td class="item-cell">$iconHtml<div class="item-main"><div class="item-title"><span class="item-name">$name</span></div>$locationHtml</div></td><td>$version</td><td>$publisher</td><td>$status</td><td class="reason">$reason</td><td>$actionHtml</td></tr>
"@
}

function Build-ReportHtml {
    param($Results, [string]$Path, [string]$CsrfToken, [string]$Nonce)
    $red = @($Results | Where-Object { $_.状态 -eq '命中黑名单' })
    $yellow = @($Results | Where-Object { $_.状态 -eq '版本变化' })
    $pending = @($Results | Where-Object { $_.状态 -eq '待定' })
    $green = @($Results | Where-Object { $_.状态 -eq '已匹配' })
    $redSoftware = @($red | Where-Object { $_.类型 -eq '软件' })
    $redPlugins = @($red | Where-Object { $_.类型 -ne '软件' })
    $tableHead = '<table><thead><tr><th class="chk-cell"></th><th>图标 / 名称</th><th>版本</th><th>发布者</th><th>状态</th><th>原因</th><th>操作</th></tr></thead><tbody>'
    $tableTail = '</tbody></table>'
    $body = ''
    $sections = @(
        [PSCustomObject]@{ Title = '&#128308; 软件核对（命中黑名单）'; Class = 'sec-red'; Items = $redSoftware; Empty = '软件部分没有发现异常。' },
        [PSCustomObject]@{ Title = '&#128308; 浏览器插件核对（命中黑名单）'; Class = 'sec-red'; Items = $redPlugins; Empty = '插件部分没有发现异常。' },
        [PSCustomObject]@{ Title = '&#128993; 版本变化（名称匹配，版本号不同）'; Class = 'sec-yellow'; Items = $yellow; Empty = '无' },
        [PSCustomObject]@{ Title = '&#128309; 待定（请用户归类）'; Class = 'sec-pending'; Items = $pending; Empty = '无' }
    )
    foreach ($section in $sections) {
        $body += "<div class='section-container'><div class='section-title'><h2 class='$($section.Class)'>$($section.Title)</h2><label class='section-select'><input type='checkbox' class='section-chk'> 全选本区</label></div>"
        if ($section.Items.Count -gt 0) { $body += $tableHead + (($section.Items | ForEach-Object { Get-RowHtml -Item $_ -NeedsAction $true }) -join "`n") + $tableTail }
        else { $body += "<p class='empty'>$(ConvertTo-HtmlEncodedText $section.Empty)</p>" }
        $body += '</div>'
    }
    $greenRows = ($green | ForEach-Object { Get-RowHtml -Item $_ -NeedsAction $false }) -join "`n"
    $body += "<div class='section-container'><details class='green-section'><summary><h2 class='sec-green'>&#128994; 已匹配（点击展开，共 $($green.Count) 项）</h2><label class='section-select green-select'><input type='checkbox' class='section-chk'> 全选本区</label></summary>$tableHead$greenRows$tableTail</details></div>"
    $summary = "<div class='summary'><span>核对时间：$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</span><span class='tag tag-red'>黑名单：$($red.Count)</span><span class='tag tag-yellow'>版本变化：$($yellow.Count)</span><span class='tag tag-pending'>待定：$($pending.Count)</span><span class='tag tag-green'>已匹配：$($green.Count)</span><span class='tag'>清单文件：$(ConvertTo-HtmlEncodedText $Path)</span><button type='button' id='openManagementButton' class='manage-button'>清单维护</button><button type='button' id='reloadButton'>重新扫描</button></div>"
    $template = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'report_template.html') -Raw -Encoding UTF8
    return $template.Replace('__SUMMARY__', $summary).Replace('__BODY__', $body).Replace('__CSRF_TOKEN__', (ConvertTo-HtmlEncodedText $CsrfToken)).Replace('__CSP_NONCE__', (ConvertTo-HtmlEncodedText $Nonce))
}

function New-RandomToken {
    param([int]$ByteCount)
    $bytes = New-Object byte[] $ByteCount
    $generator = New-Object System.Security.Cryptography.RNGCryptoServiceProvider
    try { $generator.GetBytes($bytes) } finally { $generator.Dispose() }
    return ([BitConverter]::ToString($bytes) -replace '-', '').ToLowerInvariant()
}

function Set-SecurityHeaders {
    param($Response, [string]$Nonce)
    $Response.Headers['Cache-Control'] = 'no-store, max-age=0'
    $Response.Headers['Pragma'] = 'no-cache'
    $Response.Headers['X-Content-Type-Options'] = 'nosniff'
    $Response.Headers['X-Frame-Options'] = 'DENY'
    $Response.Headers['Referrer-Policy'] = 'no-referrer'
    $Response.Headers['Content-Security-Policy'] = "default-src 'none'; script-src 'nonce-$Nonce'; style-src 'unsafe-inline'; img-src data:; connect-src 'self'; base-uri 'none'; frame-ancestors 'none'; form-action 'self'"
}

function Write-ResponseBytes {
    param($Response, [byte[]]$Bytes, [string]$ContentType, [int]$StatusCode, [string]$Nonce)
    $Response.StatusCode = $StatusCode
    $Response.ContentType = $ContentType
    $Response.ContentEncoding = [Text.Encoding]::UTF8
    $Response.ContentLength64 = $Bytes.Length
    $Response.KeepAlive = $false
    Set-SecurityHeaders -Response $Response -Nonce $Nonce
    $Response.OutputStream.Write($Bytes, 0, $Bytes.Length)
}

function Write-JsonResponse {
    param($Response, $Object, [int]$StatusCode, [string]$Nonce)
    $json = $Object | ConvertTo-Json -Compress -Depth 8
    Write-ResponseBytes -Response $Response -Bytes ([Text.Encoding]::UTF8.GetBytes($json)) -ContentType 'application/json; charset=utf-8' -StatusCode $StatusCode -Nonce $Nonce
}

function Write-HtmlResponse {
    param($Response, [string]$Html, [int]$StatusCode, [string]$Nonce)
    Write-ResponseBytes -Response $Response -Bytes ([Text.Encoding]::UTF8.GetBytes($Html)) -ContentType 'text/html; charset=utf-8' -StatusCode $StatusCode -Nonce $Nonce
}

function Test-LoopbackHost {
    param([string]$HostName)
    if ([string]::Equals($HostName, 'localhost', [StringComparison]::OrdinalIgnoreCase)) { return $true }
    $address = $null
    if ([Net.IPAddress]::TryParse($HostName, [ref]$address)) { return [Net.IPAddress]::IsLoopback($address) }
    return $false
}

function Assert-AuthorizedPostRequest {
    param($Request, [string]$CsrfToken)
    if ($Request.HttpMethod -ne 'POST') { throw '只允许 POST。' }
    if (-not (Test-LoopbackHost -HostName $Request.Url.Host)) { throw '请求主机不是本机回环地址。' }
    $contentType = [string]$Request.ContentType
    if (($contentType -split ';')[0].Trim().ToLowerInvariant() -ne 'application/json') { throw '请求 Content-Type 必须为 application/json。' }
    if ([string]$Request.Headers['X-CheckSentry-Token'] -ne $CsrfToken) { throw '安全令牌无效，请刷新页面后重试。' }
    $originText = [string]$Request.Headers['Origin']
    if (-not [string]::IsNullOrWhiteSpace($originText)) {
        $origin = $null
        if (-not [Uri]::TryCreate($originText, [UriKind]::Absolute, [ref]$origin)) { throw 'Origin 无效。' }
        if ($origin.Scheme -ne 'http' -or -not (Test-LoopbackHost -HostName $origin.Host) -or $origin.Port -ne $Request.Url.Port) { throw '拒绝跨站请求。' }
    }
    if ($Request.ContentLength64 -gt $script:MaxRequestChars) { throw '请求正文过大。' }
}

function Read-JsonRequestBody {
    param($Request)
    $reader = New-Object System.IO.StreamReader($Request.InputStream, [Text.Encoding]::UTF8, $true, 4096, $true)
    try {
        $builder = New-Object Text.StringBuilder
        $buffer = New-Object char[] 4096
        while (($read = $reader.Read($buffer, 0, $buffer.Length)) -gt 0) {
            if (($builder.Length + $read) -gt $script:MaxRequestChars) { throw '请求正文过大。' }
            $null = $builder.Append($buffer, 0, $read)
        }
        $text = $builder.ToString()
        if ([string]::IsNullOrWhiteSpace($text)) { throw '请求正文为空。' }
        return $text | ConvertFrom-Json -ErrorAction Stop
    } finally { $reader.Dispose() }
}

function Get-TargetSheetFromStatus {
    param([object]$Status)
    switch ([string]$Status) {
        '允许' { return '白名单' }
        '禁止' { return '黑名单' }
        '待定' { return '待定' }
        default { throw '状态必须是允许、禁止或待定。' }
    }
}

function Write-RuleChange {
    param([string]$Path, $Data, [string]$TargetSheet, [switch]$ExplicitClassification)
    $rule = Convert-InputToRule -Data $Data -TargetSheet $TargetSheet
    $identity = Get-RuleIdentity -ItemType $rule.类型 -ExtensionId $rule.插件ID -NamePattern $rule.软件名关键词
    $originalSheet = [string]$Data.originalSheet
    $originalId = [string]$Data.id
    $isExplicitClassification = $ExplicitClassification.IsPresent
    $operation = {
        param($package)
        if (-not [string]::IsNullOrWhiteSpace($originalId)) {
            $null = Assert-AllowedSheet -SheetName $originalSheet
            Remove-RuleByIdFromPackage -Package $package -SheetName $originalSheet -RuleId $originalId
        }
        if ($isExplicitClassification -or -not [string]::IsNullOrWhiteSpace($originalId)) {
            Remove-IdentityFromPackage -Package $package -Identity $identity
        } elseif (Test-IdentityInSheet -Package $package -SheetName $TargetSheet -Identity $identity) {
            throw "【$TargetSheet】中已存在同一对象的规则。"
        }
        return Add-RuleToPackage -Package $package -SheetName $TargetSheet -Rule $rule
    }
    return Invoke-WorkbookTransaction -Path $Path -Operation $operation -CreateUndo
}

function Show-RuleConflictWarnings {
    param([string]$Path)
    $entries = @()
    foreach ($sheetName in $script:AllowedSheets) {
        foreach ($rule in @(Import-RuleSheet -Path $Path -SheetName $sheetName)) {
            $itemType = if ([string]::IsNullOrWhiteSpace([string]$rule.'类型')) { '软件' } else { [string]$rule.'类型' }
            $entries += [PSCustomObject]@{ Identity = Get-RuleIdentity -ItemType $itemType -ExtensionId ([string]$rule.'插件ID') -NamePattern ([string]$rule.'软件名关键词'); Sheet = $sheetName }
        }
    }
    foreach ($group in @($entries | Group-Object Identity | Where-Object { $_.Count -gt 1 })) {
        $sheets = ($group.Group | Select-Object -ExpandProperty Sheet -Unique) -join '、'
        $scopeText = if (($group.Group | Select-Object -ExpandProperty Sheet -Unique).Count -gt 1) { '跨表冲突' } else { '同表重复规则' }
        Write-Host "检测到$scopeText [$($group.Name)]：$sheets，共 $($group.Count) 条。运行时仍严格按 黑名单→白名单→待定 判定，未自动删除任何规则。" -ForegroundColor Yellow
    }
}

function Start-ReportServer {
    param([int]$RequestedPort, [string]$Path, [bool]$IncludeSystem)
    if ($RequestedPort -lt 1 -or $RequestedPort -gt 65535) { throw '端口必须在 1 到 65535 之间。' }
    $csrfToken = New-RandomToken -ByteCount 32
    $nonce = New-RandomToken -ByteCount 24
    $listener = $null
    $actualPort = $RequestedPort
    for ($attempt = 0; $attempt -lt 10; $attempt++) {
        if ($actualPort -gt 65535) { break }
        try {
            $listener = New-Object Net.HttpListener
            $listener.Prefixes.Add("http://localhost:$actualPort/")
            $listener.Start()
            break
        } catch {
            if ($null -ne $listener) {
                try { $listener.Close() } catch { Write-Verbose "关闭未成功启动的监听器失败：$($_.Exception.Message)" }
                $listener = $null
            }
            $actualPort++
        }
    }
    if ($null -eq $listener -or -not $listener.IsListening) { throw '无法绑定端口。请关闭占用程序或使用 -Port 指定其他端口。' }
    $url = "http://localhost:$actualPort/"
    Write-Host "报告服务已启动：$url" -ForegroundColor Green
    if ($actualPort -ne $RequestedPort) { Write-Host "端口 $RequestedPort 已占用，已自动改用 $actualPort。" -ForegroundColor Yellow }
    Write-Host '关闭此窗口即可停止工具。' -ForegroundColor DarkGray
    try { Start-Process $url -ErrorAction Stop } catch { Write-Host "无法自动打开浏览器，请手动访问：$url" -ForegroundColor Yellow }

    try {
        while ($listener.IsListening) {
            try { $context = $listener.GetContext() } catch { if (-not $listener.IsListening) { break }; continue }
            $request = $context.Request
            $response = $context.Response
            try {
                $route = $request.Url.AbsolutePath
                if ($request.HttpMethod -eq 'GET' -and $route -eq '/manage') {
                    $template = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'management_template.html') -Raw -Encoding UTF8
                    $template = $template.Replace('__CSRF_TOKEN__', (ConvertTo-HtmlEncodedText $csrfToken)).Replace('__CSP_NONCE__', (ConvertTo-HtmlEncodedText $nonce))
                    Write-HtmlResponse -Response $response -Html $template -StatusCode 200 -Nonce $nonce
                } elseif ($request.HttpMethod -eq 'GET' -and $route -eq '/api/rules') {
                    $black = @(Convert-RulesForWeb -Rules @(Import-RuleSheet -Path $Path -SheetName '黑名单') -IncludeSystem $IncludeSystem)
                    $white = @(Convert-RulesForWeb -Rules @(Import-RuleSheet -Path $Path -SheetName '白名单') -IncludeSystem $IncludeSystem)
                    $pending = @(Convert-RulesForWeb -Rules @(Import-RuleSheet -Path $Path -SheetName '待定') -IncludeSystem $IncludeSystem)
                    Write-JsonResponse -Response $response -Object @{ ok = $true; blackRules = $black; whiteRules = $white; pendingRules = $pending; canUndo = ($null -ne $script:LastUndoSnapshot) } -StatusCode 200 -Nonce $nonce
                } elseif ($request.HttpMethod -eq 'GET' -and $route -eq '/') {
                    $inventory = Get-InventoryData -ForceRefresh -IncludeSystem:$IncludeSystem
                    $black = @(Import-RuleSheet -Path $Path -SheetName '黑名单')
                    $white = @(Import-RuleSheet -Path $Path -SheetName '白名单')
                    $pending = @(Import-RuleSheet -Path $Path -SheetName '待定')
                    $results = @()
                    $results += Get-ComplianceResult -Installed $inventory.Software -BlackRules $black -WhiteRules $white -PendingRules $pending -ItemType '软件'
                    $results += Get-ComplianceResult -Installed $inventory.Chromium -BlackRules $black -WhiteRules $white -PendingRules $pending -ItemType 'Chromium插件'
                    $results += Get-ComplianceResult -Installed $inventory.Firefox -BlackRules $black -WhiteRules $white -PendingRules $pending -ItemType 'Firefox插件'
                    $newCount = Add-NewPendingRules -Path $Path -Results $results
                    if ($newCount -gt 0) { Write-Host "本次扫描自动加入待定：$newCount 条" -ForegroundColor Yellow }
                    Write-HtmlResponse -Response $response -Html (Build-ReportHtml -Results $results -Path $Path -CsrfToken $csrfToken -Nonce $nonce) -StatusCode 200 -Nonce $nonce
                } elseif ($request.HttpMethod -eq 'GET' -and $route -eq '/favicon.ico') {
                    Write-ResponseBytes -Response $response -Bytes ([byte[]]@()) -ContentType 'image/x-icon' -StatusCode 204 -Nonce $nonce
                } elseif ($request.HttpMethod -eq 'POST') {
                    Assert-AuthorizedPostRequest -Request $request -CsrfToken $csrfToken
                    $data = Read-JsonRequestBody -Request $request
                    if ($route -eq '/api/manage/undo') {
                        Restore-UndoSnapshot -Path $Path
                        Write-JsonResponse -Response $response -Object @{ ok = $true; canUndo = $false } -StatusCode 200 -Nonce $nonce
                    } elseif ($route -eq '/api/manage/delete') {
                        $sheet = Assert-AllowedSheet -SheetName $data.sheet
                        $id = Get-LimitedText -Value $data.id -FieldName '规则ID' -MaximumLength 10 -Required
                        $operation = { param($package) Remove-RuleByIdFromPackage -Package $package -SheetName $sheet -RuleId $id }
                        $null = Invoke-WorkbookTransaction -Path $Path -Operation $operation -CreateUndo
                        Write-JsonResponse -Response $response -Object @{ ok = $true; canUndo = $true } -StatusCode 200 -Nonce $nonce
                    } elseif ($route -eq '/api/manage/save') {
                        $target = Assert-AllowedSheet -SheetName $data.targetSheet
                        $newId = Write-RuleChange -Path $Path -Data $data -TargetSheet $target
                        Write-JsonResponse -Response $response -Object @{ ok = $true; id = $newId; canUndo = $true } -StatusCode 200 -Nonce $nonce
                    } elseif ($route -eq '/api/add') {
                        $target = Get-TargetSheetFromStatus -Status $data.status
                        $data | Add-Member -NotePropertyName type -NotePropertyValue $data.itemType -Force
                        $data | Add-Member -NotePropertyName originalSheet -NotePropertyValue $data.matchedSheet -Force
                        $data | Add-Member -NotePropertyName id -NotePropertyValue $data.matchedRuleId -Force
                        $newId = Write-RuleChange -Path $Path -Data $data -TargetSheet $target -ExplicitClassification
                        Write-JsonResponse -Response $response -Object @{ ok = $true; id = $newId; canUndo = $true } -StatusCode 200 -Nonce $nonce
                    } elseif ($route -eq '/api/addBatch') {
                        $items = @($data)
                        if ($items.Count -lt 1 -or $items.Count -gt 500) { throw '批量操作必须包含 1 到 500 项。' }
                        $validated = @()
                        $seenIdentities = @{}
                        foreach ($item in $items) {
                            $target = Get-TargetSheetFromStatus -Status $item.status
                            $item | Add-Member -NotePropertyName type -NotePropertyValue $item.itemType -Force
                            $rule = Convert-InputToRule -Data $item -TargetSheet $target
                            $identity = Get-RuleIdentity -ItemType $rule.类型 -ExtensionId $rule.插件ID -NamePattern $rule.软件名关键词
                            if ($seenIdentities.ContainsKey($identity)) { throw "批量请求中存在重复对象：$identity" }
                            $seenIdentities[$identity] = $true
                            $validated += [PSCustomObject]@{ Target = $target; Rule = $rule; Identity = $identity }
                        }
                        $operation = {
                            param($package)
                            foreach ($entry in $validated) {
                                Remove-IdentityFromPackage -Package $package -Identity $entry.Identity
                                $null = Add-RuleToPackage -Package $package -SheetName $entry.Target -Rule $entry.Rule
                            }
                            return $validated.Count
                        }
                        $count = Invoke-WorkbookTransaction -Path $Path -Operation $operation -CreateUndo
                        Write-JsonResponse -Response $response -Object @{ ok = $true; successCount = $count; canUndo = $true } -StatusCode 200 -Nonce $nonce
                    } else {
                        Write-JsonResponse -Response $response -Object @{ ok = $false; error = '接口不存在。' } -StatusCode 404 -Nonce $nonce
                    }
                } else {
                    Write-JsonResponse -Response $response -Object @{ ok = $false; error = '请求方法或路径不受支持。' } -StatusCode 404 -Nonce $nonce
                }
            } catch {
                $message = $_.Exception.Message
                Write-Host "处理请求失败：$message" -ForegroundColor Red
                try {
                    if ($request.HttpMethod -eq 'GET') {
                        $safeMessage = ConvertTo-HtmlEncodedText $message
                        $html = "<!DOCTYPE html><html lang='zh-CN'><head><meta charset='UTF-8'><title>CheckSentry 错误</title><style>body{font-family:sans-serif;padding:40px}.err{color:#8e1f11;background:#fdecea;padding:16px;border-radius:8px}</style></head><body><h2>报告生成失败</h2><div class='err'>$safeMessage</div><p>请关闭 Excel，确认 list.xlsx 未损坏后刷新页面。</p></body></html>"
                        Write-HtmlResponse -Response $response -Html $html -StatusCode 500 -Nonce $nonce
                    } else {
                        Write-JsonResponse -Response $response -Object @{ ok = $false; error = $message } -StatusCode 400 -Nonce $nonce
                    }
                } catch {
                    Write-Verbose "写入错误响应失败：$($_.Exception.Message)"
                }
            } finally {
                try { $response.OutputStream.Close() } catch { Write-Verbose "关闭 HTTP 响应流失败：$($_.Exception.Message)" }
            }
        }
    } finally {
        Remove-UndoSnapshot
        if ($null -ne $listener) {
            try { $listener.Stop() } catch { Write-Verbose "停止 HTTP 监听器失败：$($_.Exception.Message)" }
            try { $listener.Close() } catch { Write-Verbose "关闭 HTTP 监听器失败：$($_.Exception.Message)" }
        }
    }
}

try {
    Initialize-ImportExcel
    . (Join-Path $PSScriptRoot 'Get-InstalledSoftware.ps1')
    . (Join-Path $PSScriptRoot 'Get-InstalledExtensions.ps1')

    if ([string]::IsNullOrWhiteSpace($ListPath)) { $ListPath = Join-Path $PSScriptRoot 'list.xlsx' }
    $ListPath = [System.IO.Path]::GetFullPath($ListPath)
    if ([System.IO.Path]::GetExtension($ListPath) -ne '.xlsx') { throw '清单路径必须以 .xlsx 结尾。' }
    $templatePath = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot 'list_template.xlsx'))
    if ([string]::Equals($ListPath, $templatePath, [StringComparison]::OrdinalIgnoreCase)) { throw '不能把 list_template.xlsx 直接作为运行清单，请使用 list.xlsx。' }
    $parentDirectory = [System.IO.Path]::GetDirectoryName($ListPath)
    if (-not (Test-Path -LiteralPath $parentDirectory -PathType Container)) { throw "清单所在目录不存在：$parentDirectory" }
    if (-not (Test-Path -LiteralPath $ListPath -PathType Leaf)) {
        if (-not (Test-Path -LiteralPath $templatePath -PathType Leaf)) { throw '找不到 list_template.xlsx。' }
        Assert-WorkbookPathSchema -Path $templatePath
        Copy-Item -LiteralPath $templatePath -Destination $ListPath -ErrorAction Stop
        Write-Host "已从空白模板创建清单：$ListPath" -ForegroundColor Yellow
    }
    Assert-WorkbookPathSchema -Path $ListPath
    Show-RuleConflictWarnings -Path $ListPath
    Start-ReportServer -RequestedPort $Port -Path $ListPath -IncludeSystem $IncludeSystemComponents.IsPresent
} catch {
    Write-Host "CheckSentry 启动失败：$($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
