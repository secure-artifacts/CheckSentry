<#
.SYNOPSIS
    CheckSentry 本地 Windows IT 合规核对工具。
    黑名单优先级最高，白名单次之，待定最低。所有清单写入均通过临时副本校验后提交。
#>

param(
    [string]$ListPath = '',
    [string]$LogPath = '',
    [int]$Port = 8787,
    [switch]$IncludeSystemComponents,
    [switch]$IncludeInactiveExtensions,
    [switch]$ScanAllUsers,
    [switch]$LibraryOnly
)

$script:Utf8ConsoleEncoding = New-Object System.Text.UTF8Encoding($false)
try { [Console]::InputEncoding = $script:Utf8ConsoleEncoding } catch {}
try { [Console]::OutputEncoding = $script:Utf8ConsoleEncoding } catch {}
$global:OutputEncoding = $script:Utf8ConsoleEncoding

$ErrorActionPreference = 'Stop'
$script:AllowedSheets = @('白名单', '待定', '黑名单')
$script:AllowedItemTypes = @('软件', 'Chromium插件', 'Firefox插件')
$script:AllowedMatchTypes = @('插件ID精确', '精确', '包含', '通配符', '正则')
$script:LastUndoSnapshot = $null
$script:InventoryCache = $null
$script:IconDataCache = @{}
$script:RuleIconCache = @{}
$script:RuleIconDataById = @{}
$script:ReportIconData = @{}
$script:IconSourceByKey = @{}
$script:IconKeyByIdentity = @{}
$script:MaxRequestChars = 1048576
$script:TakeoverUnlocked = $false
$script:InventoryScanJobs = @{}
$script:InventoryScanResults = @{}
$script:InventoryScanError = ''
$script:InventoryScanStartedAt = $null
$script:DataDirectory = $PSScriptRoot
$script:CloudSyncState = [PSCustomObject]@{
    Configured = $false
    Status = '未配置'
    UsingCache = $false
    Message = ''
    SyncedAt = ''
    WhiteCount = 0
    PendingCount = 0
    BlackCount = 0
    RuleHash = ''
}

function Initialize-ImportExcel {
    $requiredVersion = [version]'7.8.10'
    $bundledManifest = Join-Path $PSScriptRoot 'Modules\ImportExcel\7.8.10\ImportExcel.psd1'
    if (Test-Path -LiteralPath $bundledManifest -PathType Leaf) {
        try {
            $moduleRoot = [System.IO.Path]::GetDirectoryName($bundledManifest)
            $integrityPath = Join-Path $moduleRoot 'CONTENT-SHA256.json'
            if (-not (Test-Path -LiteralPath $integrityPath -PathType Leaf)) { throw '缺少模块完整性清单。' }
            $parsedIntegrityEntries = Get-Content -LiteralPath $integrityPath -Raw -Encoding UTF8 | ConvertFrom-Json -ErrorAction Stop
            if ($parsedIntegrityEntries -isnot [System.Array]) { throw '模块完整性清单格式无效。' }
            [object[]]$integrityEntries = $parsedIntegrityEntries
            if ($integrityEntries.Count -lt 1 -or $integrityEntries.Count -gt 2000) { throw '模块完整性清单数量无效。' }
            $moduleRootPrefix = [System.IO.Path]::GetFullPath($moduleRoot).TrimEnd('\', '/') + [System.IO.Path]::DirectorySeparatorChar
            foreach ($entry in $integrityEntries) {
                $candidate = [System.IO.Path]::GetFullPath((Join-Path $moduleRoot ([string]$entry.path)))
                if (-not $candidate.StartsWith($moduleRootPrefix, [System.StringComparison]::OrdinalIgnoreCase)) { throw '模块完整性清单包含越界路径。' }
                if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) { throw "模块文件缺失：$($entry.path)" }
                $actualHash = (Get-FileHash -LiteralPath $candidate -Algorithm SHA256 -ErrorAction Stop).Hash
                if (-not [string]::Equals($actualHash, [string]$entry.sha256, [System.StringComparison]::OrdinalIgnoreCase)) { throw "模块文件校验失败：$($entry.path)" }
            }
            Import-Module $bundledManifest -Force -ErrorAction Stop
            foreach ($commandName in @('Open-ExcelPackage', 'Close-ExcelPackage')) {
                if ($null -eq (Get-Command $commandName -ErrorAction SilentlyContinue)) { throw "ImportExcel 缺少命令 $commandName" }
            }
            return
        } catch {
            throw "随包提供的 ImportExcel $requiredVersion 加载失败：$($_.Exception.Message)"
        }
    }
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
    return @('id', '类型', '插件ID', '匹配方式', '软件名关键词', '版本号', '发布者', '状态/分类', '备注/原因', '添加人', '添加时间')
}

function Get-CanonicalRuleStatus {
    param([string]$SheetName)
    switch (Assert-AllowedSheet -SheetName $SheetName) {
        '黑名单' { return '命中黑名单' }
        '待定' { return '待定' }
        '白名单' { return '已匹配' }
    }
}

function Get-RuleSheetHeaderMap {
    param($Worksheet)
    $map = @{}
    if ($null -eq $Worksheet.Dimension) { return $map }
    for ($column = 1; $column -le $Worksheet.Dimension.End.Column; $column++) {
        $header = [string]$Worksheet.Cells[1, $column].Text
        if (-not [string]::IsNullOrWhiteSpace($header)) { $map[$header] = $column }
    }
    return $map
}

function Get-RuleSheetCell {
    param($Worksheet, [hashtable]$HeaderMap, [int]$Row, [string[]]$Names)
    foreach ($name in $Names) {
        if ($HeaderMap.ContainsKey($name)) { return $Worksheet.Cells[$Row, $HeaderMap[$name]] }
    }
    return $null
}

function Get-RuleSheetCellValue {
    param($Worksheet, [hashtable]$HeaderMap, [int]$Row, [string[]]$Names)
    $cell = Get-RuleSheetCell -Worksheet $Worksheet -HeaderMap $HeaderMap -Row $Row -Names $Names
    if ($null -eq $cell) { return $null }
    return $cell.Value
}

function Get-SafeHttpUrl {
    param([object]$Value)
    $text = if ($null -eq $Value) { '' } else { ([string]$Value).Trim() }
    if ([string]::IsNullOrWhiteSpace($text)) { return '' }
    $parseText = $text -replace '&amp;', '&'
    $uri = $null
    if (-not [Uri]::TryCreate($parseText, [UriKind]::Absolute, [ref]$uri)) { return '' }
    if (@('http', 'https') -notcontains $uri.Scheme.ToLowerInvariant()) { return '' }
    if (-not [string]::IsNullOrWhiteSpace($uri.UserInfo)) { return '' }
    return $uri.AbsoluteUri
}

function Get-CellHyperlinkUrl {
    param($Cell)
    if ($null -eq $Cell) { return '' }
    try {
        $hyperlink = $Cell.Hyperlink
        if ($null -eq $hyperlink) { return '' }
        $candidates = @()
        if ($hyperlink -is [Uri]) {
            $candidates += $hyperlink.AbsoluteUri
            $candidates += $hyperlink.OriginalString
        }
        foreach ($propertyName in @('ExternalUri', 'Uri', 'OriginalString', 'Address', 'Url', 'Href')) {
            $property = $hyperlink.PSObject.Properties[$propertyName]
            if ($null -ne $property -and $null -ne $property.Value) {
                $candidates += [string]$property.Value
            }
        }
        $candidates += [string]$hyperlink
        foreach ($candidate in $candidates) {
            $safeUrl = Get-SafeHttpUrl -Value $candidate
            if (-not [string]::IsNullOrWhiteSpace($safeUrl)) { return $safeUrl }
        }
    } catch { return '' }
    return ''
}

function Get-UserNoteText {
    param([object]$Value)
    $text = [string]$Value
    if ($text -eq '扫描自动加入待定' -or $text -eq '扫描到的新项目，待归类' -or $text -eq '命中黑名单' -or $text -eq '待定') { return '' }
    return $text
}

function Convert-RuleSheetToCanonical {
    param($Package, [string]$SheetName)
    $sheet = Assert-AllowedSheet -SheetName $SheetName
    $worksheet = $Package.Workbook.Worksheets[$sheet]
    $headers = Get-ExpectedHeaders -SheetName $sheet
    $headerMap = Get-RuleSheetHeaderMap -Worksheet $worksheet
    $records = @()
    if ($null -ne $worksheet.Dimension -and $worksheet.Dimension.End.Row -ge 2) {
        for ($row = 2; $row -le $worksheet.Dimension.End.Row; $row++) {
            $hasValue = $false
            for ($column = 1; $column -le $worksheet.Dimension.End.Column; $column++) {
                $cell = $worksheet.Cells[$row, $column]
                if (-not [string]::IsNullOrWhiteSpace([string]$cell.Formula)) { throw "工作表【$sheet】单元格 $($cell.Address) 含公式，无法安全迁移。" }
                if (-not [string]::IsNullOrWhiteSpace([string]$cell.Value)) { $hasValue = $true }
            }
            if (-not $hasValue) { continue }
            $noteOrReason = Get-UserNoteText (Get-RuleSheetCellValue -Worksheet $worksheet -HeaderMap $headerMap -Row $row -Names @('备注/原因', '备注', '禁止原因'))
            $record = [ordered]@{
                id = Get-RuleSheetCellValue -Worksheet $worksheet -HeaderMap $headerMap -Row $row -Names @('id')
                类型 = Get-RuleSheetCellValue -Worksheet $worksheet -HeaderMap $headerMap -Row $row -Names @('类型')
                插件ID = Get-RuleSheetCellValue -Worksheet $worksheet -HeaderMap $headerMap -Row $row -Names @('插件ID')
                匹配方式 = Get-RuleSheetCellValue -Worksheet $worksheet -HeaderMap $headerMap -Row $row -Names @('匹配方式')
                软件名关键词 = Get-RuleSheetCellValue -Worksheet $worksheet -HeaderMap $headerMap -Row $row -Names @('软件名关键词')
                版本号 = Get-RuleSheetCellValue -Worksheet $worksheet -HeaderMap $headerMap -Row $row -Names @('版本号')
                发布者 = Get-RuleSheetCellValue -Worksheet $worksheet -HeaderMap $headerMap -Row $row -Names @('发布者')
                '状态/分类' = Get-CanonicalRuleStatus -SheetName $sheet
                '备注/原因' = $noteOrReason
                '备注/原因链接' = Get-CellHyperlinkUrl -Cell (Get-RuleSheetCell -Worksheet $worksheet -HeaderMap $headerMap -Row $row -Names @('备注/原因', '备注', '禁止原因'))
                添加人 = Get-RuleSheetCellValue -Worksheet $worksheet -HeaderMap $headerMap -Row $row -Names @('添加人')
                添加时间 = Get-RuleSheetCellValue -Worksheet $worksheet -HeaderMap $headerMap -Row $row -Names @('添加时间')
            }
            $records += [PSCustomObject]$record
        }
    }
    if ($null -ne $worksheet.Dimension) { $worksheet.Cells[$worksheet.Dimension.Address].Clear() }
    for ($column = 1; $column -le $headers.Count; $column++) {
        $worksheet.Cells[1, $column].Value = $headers[$column - 1]
        $worksheet.Cells[1, $column].Style.Font.Bold = $true
    }
    $targetRow = 2
    foreach ($record in $records) {
        for ($column = 1; $column -le $headers.Count; $column++) {
            $header = $headers[$column - 1]
            $recordProperty = $record.PSObject.Properties[$header]
            $recordValue = if ($null -ne $recordProperty) { $recordProperty.Value } else { '' }
            $worksheet.Cells[$targetRow, $column].Value = $recordValue
            if ($header -eq '备注/原因') {
                $noteLink = Get-SafeHttpUrl -Value $record.'备注/原因链接'
                if (-not [string]::IsNullOrWhiteSpace($noteLink)) { $worksheet.Cells[$targetRow, $column].Hyperlink = [Uri]$noteLink }
            }
            $worksheet.Cells[$targetRow, $column].Style.Numberformat.Format = '@'
        }
        $targetRow++
    }
}

function Test-WorkbookCanonicalSchema {
    param([string]$Path)
    $package = $null
    try {
        $package = Open-ExcelPackage -Path $Path -ErrorAction Stop
        foreach ($sheetName in $script:AllowedSheets) {
            $worksheet = $package.Workbook.Worksheets[$sheetName]
            $headers = Get-ExpectedHeaders -SheetName $sheetName
            if ($null -eq $worksheet -or $null -eq $worksheet.Dimension) { return $true }
            for ($column = 1; $column -le $headers.Count; $column++) {
                if ([string]($worksheet.Cells[1, $column].Text) -ne $headers[$column - 1]) { return $true }
            }
            if ($worksheet.Dimension.End.Column -gt $headers.Count) {
                for ($column = $headers.Count + 1; $column -le $worksheet.Dimension.End.Column; $column++) {
                    for ($row = 1; $row -le $worksheet.Dimension.End.Row; $row++) {
                        if (-not [string]::IsNullOrWhiteSpace([string]($worksheet.Cells[$row, $column].Value))) { return $true }
                    }
                }
            }
        }
        return $false
    } finally {
        if ($null -ne $package) { $package.Dispose() }
    }
}

function Repair-WorkbookSchema {
    param([string]$Path)
    Test-ListFileAvailable -Path $Path
    $fingerprintBefore = Get-FileFingerprint -Path $Path
    $directory = [System.IO.Path]::GetDirectoryName($Path)
    $temporaryPath = Join-Path $directory ('.' + [System.IO.Path]::GetFileName($Path) + '.' + [guid]::NewGuid().ToString('N') + '.schema.xlsx')
    $package = $null
    try {
        Copy-Item -LiteralPath $Path -Destination $temporaryPath -ErrorAction Stop
        $package = Open-ExcelPackage -Path $temporaryPath -ErrorAction Stop
        foreach ($sheetName in $script:AllowedSheets) { Convert-RuleSheetToCanonical -Package $package -SheetName $sheetName }
        Assert-WorkbookPackageSchema -Package $package
        Close-ExcelPackage -ExcelPackage $package -ErrorAction Stop
        $package = $null
        if (-not (Test-WorkbookXmlNamespace -Path $temporaryPath)) { Normalize-WorkbookXmlNamespace -Path $temporaryPath }
        if (-not (Test-WorkbookXmlNamespace -Path $temporaryPath)) { throw "清单字段迁移后的 XML 未通过标准检查：$Path" }
        Assert-WorkbookPathSchema -Path $temporaryPath
        if ((Get-FileFingerprint -Path $Path) -ne $fingerprintBefore) { throw '清单在字段迁移期间被外部修改，已取消迁移。' }
        Set-WorkbookFileFromTemp -NewPath $temporaryPath -DestinationPath $Path
    } finally {
        if ($null -ne $package) { $package.Dispose() }
        if (Test-Path -LiteralPath $temporaryPath -PathType Leaf) { Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue }
    }
}

function Ensure-CanonicalWorkbookSchema {
    param([string]$Path)
    if (Test-WorkbookCanonicalSchema -Path $Path) {
        Write-Host "检测到旧版清单字段，正在安全统一三张规则表：$Path" -ForegroundColor Yellow
        Repair-WorkbookSchema -Path $Path
        Write-Host "清单字段已统一为状态/分类与备注/原因：$Path" -ForegroundColor Green
    }
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

function Test-WorkbookXmlNamespace {
    param([string]$Path)
    $zip = $null
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
        $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
        $worksheetEntries = @($zip.Entries | Where-Object { $_.FullName -match '^xl/worksheets/sheet\d+\.xml$' })
        if ($worksheetEntries.Count -lt 1) { return $false }
        $mainNamespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
        foreach ($entry in $worksheetEntries) {
            $stream = $null
            $reader = $null
            try {
                $stream = $entry.Open()
                $reader = New-Object System.IO.StreamReader($stream)
                $document = New-Object System.Xml.XmlDocument
                $document.PreserveWhitespace = $false
                $document.Load($reader)
                $sheetDataNodes = @($document.SelectNodes("//*[local-name()='sheetData']"))
                if ($sheetDataNodes.Count -ne 1) { return $false }
                foreach ($node in $sheetDataNodes) {
                    if ([string]$node.NamespaceURI -ne $mainNamespace) { return $false }
                }
                $invalidRows = @($document.SelectNodes("//*[local-name()='row' and namespace-uri()!='$mainNamespace']"))
                $invalidCells = @($document.SelectNodes("//*[local-name()='c' and namespace-uri()!='$mainNamespace']"))
                if ($invalidRows.Count -gt 0 -or $invalidCells.Count -gt 0) { return $false }
            } catch {
                return $false
            } finally {
                if ($null -ne $reader) { $reader.Dispose() }
                elseif ($null -ne $stream) { $stream.Dispose() }
            }
        }
        return $true
    } catch {
        return $false
    } finally {
        if ($null -ne $zip) { $zip.Dispose() }
    }
}

function Normalize-WorkbookXmlNamespace {
    param([string]$Path)
    $mainNamespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    $directory = [System.IO.Path]::GetDirectoryName($Path)
    $temporaryPath = Join-Path $directory ('.' + [System.IO.Path]::GetFileName($Path) + '.' + [guid]::NewGuid().ToString('N') + '.normalized.xlsx')
    $sourceZip = $null
    $destinationZip = $null
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
        $sourceZip = [System.IO.Compression.ZipFile]::OpenRead($Path)
        $destinationZip = [System.IO.Compression.ZipFile]::Open($temporaryPath, [System.IO.Compression.ZipArchiveMode]::Create)
        foreach ($sourceEntry in $sourceZip.Entries) {
            $destinationEntry = $destinationZip.CreateEntry($sourceEntry.FullName)
            $destinationStream = $null
            $sourceStream = $null
            try {
                $destinationStream = $destinationEntry.Open()
                $sourceStream = $sourceEntry.Open()
                if ($sourceEntry.FullName -match '^xl/worksheets/sheet\d+\.xml$') {
                    $reader = $null
                    try {
                        $utf8NoBom = New-Object -TypeName System.Text.UTF8Encoding -ArgumentList $false
                        $reader = New-Object -TypeName System.IO.StreamReader -ArgumentList @($sourceStream, $utf8NoBom, $true)
                        $xmlText = $reader.ReadToEnd()
                        # 强力清除带有 r: 属性的相关节点，避免 LoadXml 时报 r: is an undeclared prefix 错误
                        $xmlText = $xmlText -replace '(?s)<hyperlinks(?:[^>]*)>.*?</hyperlinks>', ''
                        $xmlText = $xmlText -replace '(?s)<x:hyperlinks(?:[^>]*)>.*?</x:hyperlinks>', ''
                        $xmlText = $xmlText -replace '(?s)<legacyDrawing[^>]*r:id[^>]*>', ''
                        $xmlText = $xmlText -replace '(?s)<drawing[^>]*r:id[^>]*>', ''
                        $document = New-Object System.Xml.XmlDocument
                        $document.PreserveWhitespace = $true
                        $document.LoadXml($xmlText)
                        $root = $document.DocumentElement
                        if ($null -eq $root -or $root.LocalName -ne 'worksheet') {
                            throw "工作表 XML 根节点无效：$($sourceEntry.FullName)"
                        }
                        $sheetDataNodes = @($document.SelectNodes("//*[local-name()='sheetData']"))
                        if ($sheetDataNodes.Count -gt 1) {
                            for ($index = 1; $index -lt $sheetDataNodes.Count; $index++) {
                                $node = $sheetDataNodes[$index]
                                if ($null -ne $node.ParentNode) { $node.ParentNode.RemoveChild($node) | Out-Null }
                            }
                        }
                        if ([string]$root.GetAttribute('xmlns') -ne $mainNamespace) {
                            $root.SetAttribute('xmlns', $mainNamespace)
                        }
                        $unqualifiedElements = @($document.SelectNodes("//*[namespace-uri()='']"))
                        foreach ($element in $unqualifiedElements) {
                            if ($element -eq $root -or $null -eq $element.ParentNode) { continue }
                            $replacement = $document.CreateElement($element.LocalName, $mainNamespace)
                            foreach ($attribute in @($element.Attributes)) {
                                if ([string]$attribute.NamespaceURI -eq 'http://www.w3.org/2000/xmlns/') { continue }
                                if ([string]::IsNullOrWhiteSpace([string]$attribute.NamespaceURI)) {
                                    $replacement.SetAttribute($attribute.Name, $attribute.Value)
                                } else {
                                    $replacement.SetAttribute($attribute.LocalName, $attribute.NamespaceURI, $attribute.Value)
                                }
                            }
                            while ($element.HasChildNodes) {
                                $replacement.AppendChild($element.FirstChild) | Out-Null
                            }
                            $element.ParentNode.ReplaceChild($replacement, $element) | Out-Null
                        }
                        $duplicateChildren = @{}
                        foreach ($child in @($root.ChildNodes | Where-Object { $_.NodeType -eq [System.Xml.XmlNodeType]::Element })) {
                            $key = [string]$child.LocalName
                            if ($duplicateChildren.ContainsKey($key) -and $child.ChildNodes.Count -eq 0) {
                                $child.ParentNode.RemoveChild($child) | Out-Null
                            } else {
                                $duplicateChildren[$key] = $child
                            }
                        }
                        $settings = New-Object System.Xml.XmlWriterSettings
                        $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
                        $settings.OmitXmlDeclaration = $false
                        $settings.Indent = $false
                        $writer = [System.Xml.XmlWriter]::Create($destinationStream, $settings)
                        try { $document.Save($writer) } finally { $writer.Dispose() }
                    } finally {
                        if ($null -ne $reader) { $reader.Dispose() }
                    }
                } else {
                    $sourceStream.CopyTo($destinationStream)
                }
            } finally {
                if ($null -ne $sourceStream) { $sourceStream.Dispose() }
                if ($null -ne $destinationStream) { $destinationStream.Dispose() }
            }
        }
    } finally {
        if ($null -ne $destinationZip) { $destinationZip.Dispose() }
        if ($null -ne $sourceZip) { $sourceZip.Dispose() }
    }
    try {
        Set-WorkbookFileFromTemp -NewPath $temporaryPath -DestinationPath $Path
    } finally {
        if (Test-Path -LiteralPath $temporaryPath -PathType Leaf) {
            Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Repair-WorkbookFile {
    param([string]$Path)
    Normalize-WorkbookXmlNamespace -Path $Path
    if (-not (Test-WorkbookXmlNamespace -Path $Path)) {
        throw "清单 XML 命名空间修复后仍未通过标准检查：$Path"
    }
    Assert-WorkbookPathSchema -Path $Path
}

function Ensure-WorkbookCompatibility {
    param([string]$Path)
    if (-not (Test-WorkbookXmlNamespace -Path $Path)) {
        Write-Host "检测到清单存在非标准 XLSX 工作表结构，正在安全重建：$Path" -ForegroundColor Yellow
        Repair-WorkbookFile -Path $Path
        Write-Host "清单已重建为兼容 Excel 和 LibreOffice 的标准 XLSX：$Path" -ForegroundColor Green
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
            if ($headers[$column - 1] -eq '备注/原因') {
                $value = Get-UserNoteText $value
                $ordered['备注/原因链接'] = Get-CellHyperlinkUrl -Cell $cell
            }
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

function Import-AllRuleSheets {
    param([string]$Path)
    $lastError = $null
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        $package = $null
        try {
            $package = Open-ExcelPackage -Path $Path -ErrorAction Stop
            Assert-WorkbookPackageSchema -Package $package
            $result = [ordered]@{}
            foreach ($sheetName in $script:AllowedSheets) {
                $result[$sheetName] = @(Convert-WorksheetToRules -Worksheet $package.Workbook.Worksheets[$sheetName] -SheetName $sheetName)
            }
            return [PSCustomObject]$result
        } catch {
            $lastError = $_
            if ($attempt -lt 3) { Start-Sleep -Milliseconds 350 }
        } finally {
            if ($null -ne $package) { $package.Dispose() }
        }
    }
    throw "读取规则清单失败。请确认 Excel 已关闭且文件未损坏。详细错误：$($lastError.Exception.Message)"
}

function Convert-CloudWorksheetToRules {
    param($Worksheet, [string]$SheetName)
    $headers = Get-ExpectedHeaders -SheetName $SheetName
    $cloudHeaders = @($headers | Select-Object -Skip 1)
    $rules = @()
    if ($null -eq $Worksheet.Dimension -or $Worksheet.Dimension.End.Row -lt 1 -or $Worksheet.Dimension.End.Column -lt 11) {
        throw "云端工作表【$SheetName】必须包含 B:K 十列表头。"
    }
    if ($Worksheet.Dimension.End.Row -gt 100000) { throw "云端工作表【$SheetName】超过 100000 行，已拒绝处理。" }
    if (($Worksheet.Dimension.End.Row * 10) -gt 2000000) { throw "云端工作表【$SheetName】使用范围异常大，已拒绝处理。" }
    for ($index = 0; $index -lt $cloudHeaders.Count; $index++) {
        $column = $index + 2
        $actual = [string]$Worksheet.Cells[1, $column].Text
        if ($actual -ne $cloudHeaders[$index]) {
            throw "云端工作表【$SheetName】第 $column 列（B:K）应为【$($cloudHeaders[$index])】，实际为【$actual】。"
        }
    }
    for ($row = 1; $row -le $Worksheet.Dimension.End.Row; $row++) {
        for ($column = 2; $column -le 11; $column++) {
            $cell = $Worksheet.Cells[$row, $column]
            if (-not [string]::IsNullOrWhiteSpace([string]$cell.Formula)) {
                throw "云端工作表【$SheetName】单元格 $($cell.Address) 含公式，规则模板只允许纯文本/数值。"
            }
        }
    }
    for ($row = 2; $row -le $Worksheet.Dimension.End.Row; $row++) {
        $hasRuleIdentity = $false
        $ordered = [ordered]@{}
        for ($index = 0; $index -lt $cloudHeaders.Count; $index++) {
            $column = $index + 2
            $header = $cloudHeaders[$index]
            $cell = $Worksheet.Cells[$row, $column]
            $value = $cell.Value
            if ($header -eq '备注/原因') {
                $value = Get-UserNoteText $value
                $ordered['备注/原因链接'] = Get-CellHyperlinkUrl -Cell $cell
            }
            if ($index -le 3 -and $null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) { $hasRuleIdentity = $true }
            $ordered[$header] = $value
        }
        if ($hasRuleIdentity) {
            $cloudData = [PSCustomObject]@{
                type = [string]$ordered['类型']
                extId = [string]$ordered['插件ID']
                matchType = [string]$ordered['匹配方式']
                namePattern = [string]$ordered['软件名关键词']
                version = [string]$ordered['版本号']
                publisher = [string]$ordered['发布者']
                note = [string]$ordered['备注/原因']
                noteLink = [string]$ordered['备注/原因链接']
                addedBy = [string]$ordered['添加人']
                addedTime = [string]$ordered['添加时间']
            }
            try {
                $validatedRule = Convert-InputToRule -Data $cloudData -TargetSheet $SheetName
            } catch {
                $detail = [string]$_.Exception.Message
                $fieldName = 'B:K'
                $columnName = ''
                $fieldColumns = [ordered]@{
                    '类型' = 'B'; '插件ID' = 'C'; '匹配方式' = 'D'; '软件名关键词' = 'E';
                    '版本号' = 'F'; '发布者' = 'G'; '备注/原因' = 'I'; '添加人' = 'J'; '添加时间' = 'K'
                }
                foreach ($candidateField in $fieldColumns.Keys) {
                    if ($detail -match [regex]::Escape($candidateField)) {
                        $fieldName = $candidateField
                        $columnName = [string]$fieldColumns[$candidateField]
                        break
                    }
                }
                $location = if ($columnName) { "$columnName$row（$fieldName）" } else { "第 $row 行 B:K" }
                $preview = ''
                if ($columnName) {
                    $rawValue = [string]$ordered[$fieldName]
                    $singleLine = ($rawValue -replace '[\r\n\t]+', ' ').Trim()
                    if ($singleLine.Length -gt 60) { $singleLine = $singleLine.Substring(0, 60) + '…' }
                    $preview = " 当前值长度：$($rawValue.Length)；预览：【$singleLine】。"
                }
                throw "云端工作表【$SheetName】单元格 $location 校验失败：$detail$preview"
            }
            $validatedRule.添加人 = [string]$ordered['添加人']
            $validatedRule.添加时间 = [string]$ordered['添加时间']
            $rules += $validatedRule
        }
    }
    return $rules
}

function Import-CloudRuleSheets {
    param([string]$Path)
    $lastError = $null
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        $package = $null
        try {
            $package = Open-ExcelPackage -Path $Path -ErrorAction Stop
            $result = [ordered]@{}
            foreach ($sheetName in $script:AllowedSheets) {
                $worksheet = $package.Workbook.Worksheets[$sheetName]
                if ($null -eq $worksheet) { throw "云端模板缺少工作表【$sheetName】。" }
                $result[$sheetName] = @(Convert-CloudWorksheetToRules -Worksheet $worksheet -SheetName $sheetName)
            }
            return [PSCustomObject]$result
        } catch {
            $lastError = $_
            if ($attempt -lt 3) { Start-Sleep -Milliseconds 350 }
        } finally {
            if ($null -ne $package) { $package.Dispose() }
        }
    }
    throw "读取云端规则模板失败。程序只读取【待定、白名单、黑名单】，其他工作表会自动忽略。请确认这三个工作表的 B:K 表头正确且文件可导出。详细错误：$($lastError.Exception.Message)"
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

function Enter-WorkbookMutex {
    param([string]$Path)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { $nameHash = ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes(([System.IO.Path]::GetFullPath($Path)).ToLowerInvariant()))) -replace '-', '') }
    finally { $sha.Dispose() }
    $mutex = New-Object System.Threading.Mutex($false, ('Local\CheckSentry-Workbook-' + $nameHash))
    try {
        if (-not $mutex.WaitOne([TimeSpan]::FromSeconds(15))) { throw '另一个 CheckSentry 实例正在修改清单，请稍后重试。' }
        return $mutex
    } catch {
        $mutex.Dispose()
        throw
    }
}

function Exit-WorkbookMutex {
    param($Mutex)
    if ($null -eq $Mutex) { return }
    try { $Mutex.ReleaseMutex() } catch {}
    $Mutex.Dispose()
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
    $workbookMutex = Enter-WorkbookMutex -Path $Path
    $fingerprintBefore = $null
    $pendingSnapshot = $null

    $directory = [System.IO.Path]::GetDirectoryName($Path)
    $temporaryPath = Join-Path $directory ('.' + [System.IO.Path]::GetFileName($Path) + '.' + [guid]::NewGuid().ToString('N') + '.tmp.xlsx')
    $package = $null
    try {
        Test-ListFileAvailable -Path $Path
        $fingerprintBefore = Get-FileFingerprint -Path $Path
        if ($CreateUndo) { $pendingSnapshot = New-UndoSnapshotFile -Path $Path }
        Copy-Item -LiteralPath $Path -Destination $temporaryPath -ErrorAction Stop
        $package = Open-ExcelPackage -Path $temporaryPath -ErrorAction Stop
        Assert-WorkbookPackageSchema -Package $package
        $result = & $Operation $package
        Assert-WorkbookPackageSchema -Package $package
        Close-ExcelPackage -ExcelPackage $package -ErrorAction Stop
        $package = $null
        if (-not (Test-WorkbookXmlNamespace -Path $temporaryPath)) {
            Normalize-WorkbookXmlNamespace -Path $temporaryPath
        }
        if (-not (Test-WorkbookXmlNamespace -Path $temporaryPath)) {
            throw "事务保存后的清单仍未通过标准 XLSX 命名空间检查：$temporaryPath"
        }
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
        Exit-WorkbookMutex -Mutex $workbookMutex
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

function Get-TakeoverSettingsPath {
    return (Join-Path $script:DataDirectory 'CheckSentry.settings.json')
}

function New-PasswordRecord {
    param([string]$Password)
    if ([string]::IsNullOrWhiteSpace($Password)) { throw '密码不能为空。' }
    $saltBytes = New-Object byte[] 16
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try { $rng.GetBytes($saltBytes) } finally { $rng.Dispose() }
    $salt = [Convert]::ToBase64String($saltBytes)
    $iterations = 120000
    $kdf = New-Object -TypeName System.Security.Cryptography.Rfc2898DeriveBytes -ArgumentList @($Password, $saltBytes, $iterations)
    try { $hash = [Convert]::ToBase64String($kdf.GetBytes(32)) } finally { $kdf.Dispose() }
    return [ordered]@{ salt = $salt; hash = $hash; iterations = $iterations }
}

function Test-PasswordRecordShape {
    param($Record)
    try {
        $saltBytes = [Convert]::FromBase64String([string]$Record.salt)
        $hashBytes = [Convert]::FromBase64String([string]$Record.hash)
        $iterations = [int]$Record.iterations
        return ($saltBytes.Length -ge 16 -and $hashBytes.Length -eq 32 -and $iterations -ge 10000 -and $iterations -le 1000000)
    } catch { return $false }
}

function Test-PasswordRecord {
    param([string]$Password, $Record)
    if ($null -eq $Record -or [string]::IsNullOrWhiteSpace($Password) -or -not (Test-PasswordRecordShape -Record $Record)) { return $false }
    try {
        $saltBytes = [Convert]::FromBase64String([string]$Record.salt)
        $expected = [Convert]::FromBase64String([string]$Record.hash)
        $iterations = [int]$Record.iterations
        $kdf = New-Object -TypeName System.Security.Cryptography.Rfc2898DeriveBytes -ArgumentList @($Password, $saltBytes, $iterations)
        try { $actual = $kdf.GetBytes($expected.Length) } finally { $kdf.Dispose() }
        $difference = 0
        for ($index = 0; $index -lt $expected.Length; $index++) { $difference = $difference -bor ($actual[$index] -bxor $expected[$index]) }
        return ($difference -eq 0)
    } catch { return $false }
}

function Read-TakeoverSettings {
    $path = Get-TakeoverSettingsPath
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return $null }
    try {
        $fileInfo = Get-Item -LiteralPath $path -ErrorAction Stop
        if ($fileInfo.Length -gt 16384) { throw '配置文件过大。' }
        $raw = Get-Content -LiteralPath $path -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) { throw '配置文件为空。' }
        $settings = $raw | ConvertFrom-Json
        if ([int]$settings.version -ne 1 -or $null -eq $settings.super -or $null -eq $settings.usage) { throw '配置版本或密码记录缺失。' }
        if (-not (Test-PasswordRecordShape -Record $settings.super) -or -not (Test-PasswordRecordShape -Record $settings.usage)) { throw '密码记录无效。' }
        return $settings
    } catch { throw '接管密码配置文件无效，请删除 CheckSentry.settings.json 后重新设置。' }
}

function Write-TakeoverSettings {
    param($Settings)
    $path = Get-TakeoverSettingsPath
    $directory = [System.IO.Path]::GetDirectoryName($path)
    $temporaryPath = Join-Path $directory ('.CheckSentry.settings.json.' + [guid]::NewGuid().ToString('N') + '.tmp')
    try {
        $json = $Settings | ConvertTo-Json -Depth 5
        $utf8NoBom = New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false)
        [System.IO.File]::WriteAllText($temporaryPath, $json, $utf8NoBom)
        if (Test-Path -LiteralPath $path -PathType Leaf) {
            $backupPath = $path + '.' + [guid]::NewGuid().ToString('N') + '.bak'
            try {
                [System.IO.File]::Replace($temporaryPath, $path, $backupPath, $true)
            } finally {
                if (Test-Path -LiteralPath $backupPath -PathType Leaf) { Remove-Item -LiteralPath $backupPath -Force -ErrorAction SilentlyContinue }
            }
        } else {
            [System.IO.File]::Move($temporaryPath, $path)
        }
    } finally {
        if (Test-Path -LiteralPath $temporaryPath -PathType Leaf) { Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue }
    }
}

function Set-TakeoverPasswords {
    param($Data)
    $existing = Read-TakeoverSettings
    if ($null -eq $existing) {
        $superPassword = Get-LimitedText -Value $Data.superPassword -FieldName '超级密码' -MaximumLength 256 -Required
        $usagePassword = Get-LimitedText -Value $Data.usagePassword -FieldName '使用密码' -MaximumLength 256 -Required
        $settings = [ordered]@{ version = 1; super = (New-PasswordRecord -Password $superPassword); usage = (New-PasswordRecord -Password $usagePassword) }
    } else {
        $currentSuperPassword = Get-LimitedText -Value $Data.currentSuperPassword -FieldName '当前超级密码' -MaximumLength 256 -Required
        if (-not (Test-PasswordRecord -Password $currentSuperPassword -Record $existing.super)) { throw '超级密码不正确。' }
        $newUsagePassword = Get-LimitedText -Value $Data.usagePassword -FieldName '使用密码' -MaximumLength 256 -Required
        $newSuperPassword = Get-LimitedText -Value $Data.newSuperPassword -FieldName '新超级密码' -MaximumLength 256
        $superRecord = if ([string]::IsNullOrWhiteSpace($newSuperPassword)) { $existing.super } else { New-PasswordRecord -Password $newSuperPassword }
        $settings = [ordered]@{ version = 1; super = $superRecord; usage = (New-PasswordRecord -Password $newUsagePassword) }
    }
    Write-TakeoverSettings -Settings $settings
}

function Get-CloudCacheDirectory {
    return (Join-Path $script:DataDirectory 'CloudCache')
}

function Get-CloudSnapshotPath {
    return (Join-Path (Get-CloudCacheDirectory) 'last-good.xlsx')
}

function Get-CloudStatePath {
    return (Join-Path (Get-CloudCacheDirectory) 'sync-state.json')
}

function Read-CloudSyncMetadata {
    $path = Get-CloudStatePath
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return $null }
    try {
        $item = Get-Item -LiteralPath $path -ErrorAction Stop
        if ($item.Length -gt 65536) { throw '云端同步状态文件过大。' }
        return (Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json -ErrorAction Stop)
    } catch {
        Write-Verbose "读取云端同步状态失败：$($_.Exception.Message)"
        return $null
    }
}

function Write-CloudSyncMetadata {
    param($Metadata)
    $directory = Get-CloudCacheDirectory
    if (-not (Test-Path -LiteralPath $directory -PathType Container)) { New-Item -ItemType Directory -Path $directory -Force | Out-Null }
    $path = Get-CloudStatePath
    $temporaryPath = Join-Path $directory ('.sync-state.' + [guid]::NewGuid().ToString('N') + '.tmp')
    try {
        $json = $Metadata | ConvertTo-Json -Depth 5
        $utf8NoBom = New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false)
        [System.IO.File]::WriteAllText($temporaryPath, $json, $utf8NoBom)
        if (Test-Path -LiteralPath $path -PathType Leaf) {
            $backupPath = Join-Path $directory ('.sync-state.' + [guid]::NewGuid().ToString('N') + '.bak')
            try { [System.IO.File]::Replace($temporaryPath, $path, $backupPath, $true) }
            finally { if (Test-Path -LiteralPath $backupPath -PathType Leaf) { Remove-Item -LiteralPath $backupPath -Force -ErrorAction SilentlyContinue } }
        } else {
            [System.IO.File]::Move($temporaryPath, $path)
        }
    } finally {
        if (Test-Path -LiteralPath $temporaryPath -PathType Leaf) { Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue }
    }
}

function Save-CloudSnapshot {
    param([string]$SourcePath)
    $directory = Get-CloudCacheDirectory
    if (-not (Test-Path -LiteralPath $directory -PathType Container)) { New-Item -ItemType Directory -Path $directory -Force | Out-Null }
    $destination = Get-CloudSnapshotPath
    $temporaryPath = Join-Path $directory ('.last-good.' + [guid]::NewGuid().ToString('N') + '.tmp.xlsx')
    try {
        Copy-Item -LiteralPath $SourcePath -Destination $temporaryPath -ErrorAction Stop
        if (Test-Path -LiteralPath $destination -PathType Leaf) {
            $backupPath = Join-Path $directory ('.last-good.' + [guid]::NewGuid().ToString('N') + '.bak.xlsx')
            try { [System.IO.File]::Replace($temporaryPath, $destination, $backupPath, $true) }
            finally { if (Test-Path -LiteralPath $backupPath -PathType Leaf) { Remove-Item -LiteralPath $backupPath -Force -ErrorAction SilentlyContinue } }
        } else {
            [System.IO.File]::Move($temporaryPath, $destination)
        }
    } finally {
        if (Test-Path -LiteralPath $temporaryPath -PathType Leaf) { Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue }
    }
}

function Get-CloudSettingsPath {
    return (Join-Path $script:DataDirectory 'CheckSentry.cloud.json')
}

function Read-CloudSettings {
    $path = Get-CloudSettingsPath
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return $null }
    try {
        $fileInfo = Get-Item -LiteralPath $path -ErrorAction Stop
        if ($fileInfo.Length -gt 16384) { throw '配置文件过大。' }
        $raw = Get-Content -LiteralPath $path -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) { throw '配置文件为空。' }
        $settings = $raw | ConvertFrom-Json -ErrorAction Stop
        if ([int]$settings.version -ne 1) { throw '配置版本无效。' }
        $null = Get-GoogleSheetsExportUrl -Url ([string]$settings.url)
        return $settings
    } catch { throw '云端清单配置文件无效，请在维护页重新设置云端链接。' }
}

function Write-CloudSettings {
    param([string]$Url)
    $exportUrl = Get-GoogleSheetsExportUrl -Url $Url
    $path = Get-CloudSettingsPath
    $directory = [System.IO.Path]::GetDirectoryName($path)
    $temporaryPath = Join-Path $directory ('.CheckSentry.cloud.json.' + [guid]::NewGuid().ToString('N') + '.tmp')
    $backupPath = $null
    try {
        $settings = [ordered]@{ version = 1; url = $Url.Trim() }
        $json = $settings | ConvertTo-Json -Depth 3
        $utf8NoBom = New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false)
        [System.IO.File]::WriteAllText($temporaryPath, $json, $utf8NoBom)
        if (Test-Path -LiteralPath $path -PathType Leaf) {
            $backupPath = $path + '.' + [guid]::NewGuid().ToString('N') + '.bak'
            try { [System.IO.File]::Replace($temporaryPath, $path, $backupPath, $true) }
            finally { if ($backupPath -and (Test-Path -LiteralPath $backupPath -PathType Leaf)) { Remove-Item -LiteralPath $backupPath -Force -ErrorAction SilentlyContinue } }
        } else {
            [System.IO.File]::Move($temporaryPath, $path)
        }
    } finally {
        if (Test-Path -LiteralPath $temporaryPath -PathType Leaf) { Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue }
    }
    return $exportUrl
}

function Get-GoogleSheetsExportUrl {
    param([string]$Url)
    $text = if ($null -eq $Url) { '' } else { $Url.Trim() }
    if ([string]::IsNullOrWhiteSpace($text) -or $text.Length -gt 2048) { throw 'Google Sheets 链接不能为空且不能超过 2048 个字符。' }
    $uri = $null
    if (-not [Uri]::TryCreate($text, [UriKind]::Absolute, [ref]$uri)) { throw 'Google Sheets 链接格式无效。' }
    if ($uri.Scheme -ne 'https' -or $uri.Host -ne 'docs.google.com') { throw '只允许使用 https://docs.google.com/spreadsheets/d/... 链接。' }
    $match = [regex]::Match($uri.AbsolutePath, '^/spreadsheets/d/([A-Za-z0-9_-]{10,200})(?:/|$)')
    if (-not $match.Success) { throw '链接不是有效的 Google Sheets 表格链接。' }
    $spreadsheetId = $match.Groups[1].Value
    return ('https://docs.google.com/spreadsheets/d/{0}/export?format=xlsx' -f $spreadsheetId)
}

function Assert-XlsxArchiveComplete {
    param([string]$Path)
    Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    $stream = $null
    $archive = $null
    try {
        $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
        $archive = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Read, $false)
        $requiredEntries = @('[Content_Types].xml', 'xl/workbook.xml')
        foreach ($required in $requiredEntries) {
            if ($null -eq $archive.GetEntry($required)) { throw "XLSX 缺少必要文件：$required" }
        }
        $expandedBytes = [int64]0
        $buffer = New-Object byte[] 65536
        foreach ($entry in $archive.Entries) {
            $expandedBytes += [int64]$entry.Length
            if ($expandedBytes -gt 104857600) { throw 'XLSX 解压后超过 100 MB 安全限制。' }
            $entryStream = $null
            try {
                $entryStream = $entry.Open()
                while ($entryStream.Read($buffer, 0, $buffer.Length) -gt 0) {}
            } finally {
                if ($null -ne $entryStream) { $entryStream.Dispose() }
            }
        }
    } catch {
        throw "云端 XLSX 完整性检查失败：$($_.Exception.Message)"
    } finally {
        if ($null -ne $archive) { $archive.Dispose() }
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

function Download-GoogleSheetsWorkbookOnce {
    param([string]$Url)
    $exportUrl = (Get-GoogleSheetsExportUrl -Url $Url) + '&checksentry=' + [guid]::NewGuid().ToString('N')
    $temporaryPath = Join-Path ([System.IO.Path]::GetTempPath()) ('CheckSentry-cloud-' + [guid]::NewGuid().ToString('N') + '.xlsx')
    $maximumBytes = [int64]10485760
    $response = $null
    $inputStream = $null
    $outputStream = $null
    $success = $false
    try {
        $request = [System.Net.HttpWebRequest]::Create($exportUrl)
        $request.Method = 'GET'
        $request.Timeout = 30000
        $request.ReadWriteTimeout = 30000
        $request.AllowAutoRedirect = $true
        $request.MaximumAutomaticRedirections = 5
        $request.UserAgent = 'CheckSentry/1.0.5'
        $response = $request.GetResponse()
        $effectiveUri = $response.ResponseUri
        if ($null -eq $effectiveUri -or ($effectiveUri.Host -ne 'docs.google.com' -and -not $effectiveUri.Host.EndsWith('.googleusercontent.com', [StringComparison]::OrdinalIgnoreCase))) {
            throw '云端下载被重定向到不受允许的主机。'
        }
        if ($response.ContentLength -gt $maximumBytes) { throw '云端 XLSX 文件超过 10 MB 限制。' }
        $inputStream = $response.GetResponseStream()
        $outputStream = [System.IO.File]::Open($temporaryPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        $buffer = New-Object byte[] 65536
        $totalBytes = [int64]0
        while (($read = $inputStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
            $totalBytes += $read
            if ($totalBytes -gt $maximumBytes) { throw '云端 XLSX 文件超过 10 MB 限制。' }
            $outputStream.Write($buffer, 0, $read)
        }
        $outputStream.Flush()
        if ($totalBytes -le 0) { throw '云端 XLSX 文件为空。' }
        $outputStream.Dispose()
        $outputStream = $null
        $inputStream.Dispose()
        $inputStream = $null
        $response.Close()
        $response = $null
        $stream = [System.IO.File]::OpenRead($temporaryPath)
        try {
            $header = New-Object byte[] 2
            $null = $stream.Read($header, 0, 2)
        } finally { $stream.Dispose() }
        if ($header[0] -ne 80 -or $header[1] -ne 75) { throw '云端链接没有返回有效的 XLSX 文件，可能需要登录或调整 Google Sheets 分享权限。' }
        Assert-XlsxArchiveComplete -Path $temporaryPath
        $success = $true
        return $temporaryPath
    } catch {
        throw "下载 Google Sheets 规则模板失败：$($_.Exception.Message)"
    } finally {
        if ($null -ne $outputStream) { $outputStream.Dispose() }
        if ($null -ne $inputStream) { $inputStream.Dispose() }
        if ($null -ne $response) { $response.Close() }
        if (-not $success -and (Test-Path -LiteralPath $temporaryPath -PathType Leaf)) {
            Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Download-GoogleSheetsWorkbook {
    param([string]$Url, [int]$MaximumAttempts = 3)
    $lastError = $null
    for ($attempt = 1; $attempt -le $MaximumAttempts; $attempt++) {
        try {
            return (Download-GoogleSheetsWorkbookOnce -Url $Url)
        } catch {
            $lastError = $_
            if ($attempt -lt $MaximumAttempts) {
                Start-Sleep -Milliseconds ([int](500 * [math]::Pow(2, $attempt - 1)))
            }
        }
    }
    throw "云端清单下载重试 $MaximumAttempts 次仍失败：$($lastError.Exception.Message)"
}

function Get-RuleDataSignature {
    param($Rule, [string]$SheetName)
    $data = [ordered]@{
        sheet = $SheetName
        类型 = [string]$Rule.'类型'
        插件ID = [string]$Rule.'插件ID'
        匹配方式 = [string]$Rule.'匹配方式'
        软件名关键词 = [string]$Rule.'软件名关键词'
        版本号 = [string]$Rule.'版本号'
        发布者 = [string]$Rule.'发布者'
        '备注/原因' = [string]$Rule.'备注/原因'
        '备注/原因链接' = [string]$Rule.'备注/原因链接'
        添加人 = [string]$Rule.'添加人'
        添加时间 = [string]$Rule.'添加时间'
    }
    return ($data | ConvertTo-Json -Compress -Depth 3)
}

function Get-RuleSetFingerprint {
    param($AllRules)
    $signatures = @()
    foreach ($sheetName in $script:AllowedSheets) {
        foreach ($rule in @($AllRules.$sheetName)) { $signatures += Get-RuleDataSignature -Rule $rule -SheetName $sheetName }
    }
    return (($signatures | Sort-Object) -join "`n")
}

function Get-TextSha256 {
    param([string]$Text)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes([string]$Text)
        return (([BitConverter]::ToString($sha.ComputeHash($bytes))) -replace '-', '')
    } finally { $sha.Dispose() }
}

function Get-CloudRuleStats {
    param($AllRules)
    $whiteCount = @($AllRules.'白名单').Count
    $pendingCount = @($AllRules.'待定').Count
    $blackCount = @($AllRules.'黑名单').Count
    return [PSCustomObject]@{
        WhiteCount = $whiteCount
        PendingCount = $pendingCount
        BlackCount = $blackCount
        TotalCount = $whiteCount + $pendingCount + $blackCount
    }
}

function New-CloudCandidateFromPath {
    param([string]$Path, [string]$Source)
    try { Normalize-WorkbookXmlNamespace -Path $Path } catch { Write-Verbose "云端模板命名空间修复失败，将尝试直接读取：$($_.Exception.Message)" }
    Assert-XlsxArchiveComplete -Path $Path
    $rules = Import-CloudRuleSheets -Path $Path
    $stats = Get-CloudRuleStats -AllRules $rules
    if ($stats.TotalCount -lt 1) { throw '云端三张规则表没有任何规则，已拒绝启用空白云端清单。' }
    $hash = Get-TextSha256 -Text (Get-RuleSetFingerprint -AllRules $rules)
    return [PSCustomObject]@{ Path = $Path; Source = $Source; Rules = $rules; Stats = $stats; RuleHash = $hash }
}

function Get-ValidatedRemoteCloudCandidate {
    param([string]$Url)
    $firstPath = $null
    $secondPath = $null
    $returnPath = $null
    try {
        $firstPath = Download-GoogleSheetsWorkbook -Url $Url
        $firstCandidate = New-CloudCandidateFromPath -Path $firstPath -Source 'remote'
        $metadata = Read-CloudSyncMetadata
        $requiresConfirmation = ($null -eq $metadata -or [string]$metadata.url -ne $Url.Trim() -or [string]$metadata.ruleHash -ne $firstCandidate.RuleHash)
        if (-not $requiresConfirmation) {
            $returnPath = $firstPath
            return $firstCandidate
        }

        Start-Sleep -Milliseconds 750
        $secondPath = Download-GoogleSheetsWorkbook -Url $Url
        $secondCandidate = New-CloudCandidateFromPath -Path $secondPath -Source 'remote'
        if ($firstCandidate.RuleHash -ne $secondCandidate.RuleHash) {
            throw '连续两次下载的云端规则不一致，可能正在编辑或导出尚未稳定，本次未启用。'
        }
        $returnPath = $secondPath
        return $secondCandidate
    } finally {
        foreach ($candidatePath in @($firstPath, $secondPath)) {
            if ($candidatePath -and $candidatePath -ne $returnPath -and (Test-Path -LiteralPath $candidatePath -PathType Leaf)) {
                Remove-Item -LiteralPath $candidatePath -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

function Get-ValidatedCachedCloudCandidate {
    param([string]$Url)
    $metadata = Read-CloudSyncMetadata
    $snapshot = Get-CloudSnapshotPath
    if ($null -eq $metadata -or [string]$metadata.url -ne $Url.Trim()) { throw '没有与当前链接对应的云端成功快照。' }
    if (-not (Test-Path -LiteralPath $snapshot -PathType Leaf)) { throw '云端成功快照文件不存在。' }
    $temporaryPath = Join-Path ([System.IO.Path]::GetTempPath()) ('CheckSentry-cloud-cache-' + [guid]::NewGuid().ToString('N') + '.xlsx')
    try {
        Copy-Item -LiteralPath $snapshot -Destination $temporaryPath -ErrorAction Stop
        $candidate = New-CloudCandidateFromPath -Path $temporaryPath -Source 'cache'
        if ($candidate.RuleHash -ne [string]$metadata.ruleHash) { throw '云端成功快照与同步状态指纹不一致。' }
        return $candidate
    } catch {
        if (Test-Path -LiteralPath $temporaryPath -PathType Leaf) { Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue }
        throw
    }
}

function Test-CloudRulesApplied {
    param($CloudEntries, $LocalRules)
    $localByKey = @{}
    $localLocations = @{}
    foreach ($sheetName in $script:AllowedSheets) {
        foreach ($rule in @($LocalRules.$sheetName)) {
            $itemType = if ([string]::IsNullOrWhiteSpace([string]$rule.'类型')) { '软件' } else { [string]$rule.'类型' }
            $identity = Get-RuleIdentity -ItemType $itemType -ExtensionId ([string]$rule.'插件ID') -NamePattern ([string]$rule.'软件名关键词') -Publisher ([string]$rule.'发布者') -Version ([string]$rule.'版本号')
            $key = $sheetName + '|' + $identity
            if ($localByKey.ContainsKey($key)) { $localByKey[$key] = $null }
            else { $localByKey[$key] = $rule }
            if (-not $localLocations.ContainsKey($identity)) { $localLocations[$identity] = @() }
            $localLocations[$identity] += $key
        }
    }
    foreach ($entry in $CloudEntries) {
        $key = $entry.Sheet + '|' + $entry.Identity
        if (-not $localByKey.ContainsKey($key) -or $null -eq $localByKey[$key]) { return $false }
        if ($localLocations[$entry.Identity].Count -ne 1) { return $false }
        if ((Get-RuleDataSignature -Rule $entry.Rule -SheetName $entry.Sheet) -ne (Get-RuleDataSignature -Rule $localByKey[$key] -SheetName $entry.Sheet)) { return $false }
    }
    return $true
}

function Set-CloudSyncRuntimeState {
    param([bool]$Configured, [string]$Status, [bool]$UsingCache, [string]$Message, $Stats, [string]$RuleHash, [string]$SyncedAt)
    $script:CloudSyncState = [PSCustomObject]@{
        Configured = $Configured
        Status = $Status
        UsingCache = $UsingCache
        Message = $Message
        SyncedAt = $SyncedAt
        WhiteCount = if ($null -eq $Stats) { 0 } else { [int]$Stats.WhiteCount }
        PendingCount = if ($null -eq $Stats) { 0 } else { [int]$Stats.PendingCount }
        BlackCount = if ($null -eq $Stats) { 0 } else { [int]$Stats.BlackCount }
        RuleHash = $RuleHash
    }
}

function Sync-CloudWorkbook {
    param([string]$Path, [switch]$CreateUndo, [string]$Url = '', [switch]$RequireRemote)
    $settings = if ([string]::IsNullOrWhiteSpace($Url)) { Read-CloudSettings } else { [PSCustomObject]@{ url = $Url.Trim() } }
    if ($null -eq $settings) {
        Set-CloudSyncRuntimeState -Configured $false -Status '未配置' -UsingCache $false -Message '' -Stats $null -RuleHash '' -SyncedAt ''
        return [PSCustomObject]@{ Configured = $false; Changed = $false; Count = 0; UsingCache = $false }
    }
    $effectiveUrl = [string]$settings.url
    $candidate = $null
    $remoteError = ''
    try {
        try {
            $candidate = Get-ValidatedRemoteCloudCandidate -Url $effectiveUrl
        } catch {
            $remoteError = $_.Exception.Message
            if ($RequireRemote) { throw }
            try {
                $candidate = Get-ValidatedCachedCloudCandidate -Url $effectiveUrl
            } catch {
                throw "远程云端规则不可用：$remoteError；最后成功快照也不可用：$($_.Exception.Message)"
            }
        }
        $cloudRules = $candidate.Rules
        $localRules = Import-AllRuleSheets -Path $Path
        $previousMetadata = Read-CloudSyncMetadata
        $cloudEntries = @()
        $identitySet = @{}
        foreach ($sheetName in $script:AllowedSheets) {
            foreach ($rule in @($cloudRules.$sheetName)) {
                $identity = Get-RuleIdentity -ItemType $rule.'类型' -ExtensionId ([string]$rule.'插件ID') -NamePattern ([string]$rule.'软件名关键词') -Publisher ([string]$rule.'发布者') -Version ([string]$rule.'版本号')
                if ($identitySet.ContainsKey($identity)) { throw "云端模板存在重复对象或跨表冲突：$identity" }
                $identitySet[$identity] = $true
                $effectiveRule = [PSCustomObject]([ordered]@{
                    类型 = [string]$rule.'类型'
                    插件ID = [string]$rule.'插件ID'
                    匹配方式 = [string]$rule.'匹配方式'
                    软件名关键词 = [string]$rule.'软件名关键词'
                    版本号 = [string]$rule.'版本号'
                    发布者 = [string]$rule.'发布者'
                    '状态/分类' = Get-CanonicalRuleStatus -SheetName $sheetName
                    '备注/原因' = [string]$rule.'备注/原因'
                    '备注/原因链接' = [string]$rule.'备注/原因链接'
                    添加人 = [string]$rule.'添加人'
                    添加时间 = [string]$rule.'添加时间'
                })
                $cloudEntries += [PSCustomObject]@{ Sheet = $sheetName; Rule = $effectiveRule; Identity = $identity }
            }
        }
        $removalIdentitySet = @{}
        foreach ($identity in @($identitySet.Keys)) { $removalIdentitySet[[string]$identity] = $true }
        if ($null -ne $previousMetadata -and
            [string]$previousMetadata.url -eq $effectiveUrl.Trim() -and
            $null -ne $previousMetadata.cloudIdentities) {
            foreach ($identity in @($previousMetadata.cloudIdentities)) {
                if (-not [string]::IsNullOrWhiteSpace([string]$identity)) {
                    $removalIdentitySet[[string]$identity] = $true
                }
            }
        }
        $hasRetiredCloudIdentity = $false
        foreach ($identity in @($removalIdentitySet.Keys)) {
            if (-not $identitySet.ContainsKey([string]$identity)) { $hasRetiredCloudIdentity = $true; break }
        }
        $changed = -not (Test-CloudRulesApplied -CloudEntries $cloudEntries -LocalRules $localRules)
        if (-not $changed -and $hasRetiredCloudIdentity) { $changed = $true }
        if ($changed) {
            $operation = {
                param($package)
                Remove-IdentitiesFromPackage -Package $package -IdentitySet $removalIdentitySet
                $states = @{}
                foreach ($entry in $cloudEntries) {
                    if (-not $states.ContainsKey($entry.Sheet)) { $states[$entry.Sheet] = Get-WorksheetAppendState -Package $package -SheetName $entry.Sheet }
                    $null = Add-RuleToPackageWithState -State $states[$entry.Sheet] -Rule $entry.Rule
                }
                return $cloudEntries.Count
            }
            $null = Invoke-WorkbookTransaction -Path $Path -Operation $operation -CreateUndo:$CreateUndo.IsPresent
        }
        $verifiedRules = Import-AllRuleSheets -Path $Path
        if (-not (Test-CloudRulesApplied -CloudEntries $cloudEntries -LocalRules $verifiedRules)) {
            throw '云端规则写入后复核失败，本地清单与已验证云端规则不一致。'
        }

        $usingCache = $candidate.Source -eq 'cache'
        $syncedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        if (-not $usingCache) {
            Save-CloudSnapshot -SourcePath $candidate.Path
            Write-CloudSyncMetadata -Metadata ([ordered]@{
                version = 2
                url = $effectiveUrl.Trim()
                syncedAt = $syncedAt
                ruleHash = $candidate.RuleHash
                cloudIdentities = @($identitySet.Keys | Sort-Object)
                whiteCount = $candidate.Stats.WhiteCount
                pendingCount = $candidate.Stats.PendingCount
                blackCount = $candidate.Stats.BlackCount
            })
        } else {
            $metadata = Read-CloudSyncMetadata
            if ($null -ne $metadata -and -not [string]::IsNullOrWhiteSpace([string]$metadata.syncedAt)) { $syncedAt = [string]$metadata.syncedAt }
        }
        $message = if ($usingCache) { '云端暂时不可用，正在使用上次完整验证的规则。远程错误：' + $remoteError } else { '云端规则已完整验证。' }
        $runtimeStatus = if ($usingCache) { '缓存' } else { '已验证' }
        Set-CloudSyncRuntimeState -Configured $true -Status $runtimeStatus -UsingCache $usingCache -Message $message -Stats $candidate.Stats -RuleHash $candidate.RuleHash -SyncedAt $syncedAt
        return [PSCustomObject]@{
            Configured = $true
            Changed = $changed
            Count = $cloudEntries.Count
            UsingCache = $usingCache
            Message = $message
            SyncedAt = $syncedAt
            WhiteCount = $candidate.Stats.WhiteCount
            PendingCount = $candidate.Stats.PendingCount
            BlackCount = $candidate.Stats.BlackCount
            RuleHash = $candidate.RuleHash
        }
    } catch {
        Set-CloudSyncRuntimeState -Configured $true -Status '失败' -UsingCache $false -Message $_.Exception.Message -Stats $null -RuleHash '' -SyncedAt ''
        throw
    } finally {
        if ($null -ne $candidate -and $candidate.Path -and (Test-Path -LiteralPath $candidate.Path -PathType Leaf)) {
            Remove-Item -LiteralPath $candidate.Path -Force -ErrorAction SilentlyContinue
        }
    }
}

function Get-RuleMetadataByIdFromPackage {
    param($Package, [string]$SheetName, [string]$RuleId)
    $sheet = Assert-AllowedSheet -SheetName $SheetName
    if ($RuleId -notmatch '^\d{1,10}$') { throw '规则 ID 无效。' }
    $worksheet = $Package.Workbook.Worksheets[$sheet]
    if ($null -eq $worksheet.Dimension) { throw "工作表【$sheet】中不存在规则 $RuleId。" }
    $headers = Get-ExpectedHeaders -SheetName $sheet
    $idColumn = [array]::IndexOf($headers, 'id') + 1
    $addedByColumn = [array]::IndexOf($headers, '添加人') + 1
    $addedTimeColumn = [array]::IndexOf($headers, '添加时间') + 1
    for ($row = 2; $row -le $worksheet.Dimension.End.Row; $row++) {
        if ([string]($worksheet.Cells[$row, $idColumn].Value) -eq $RuleId) {
            return [PSCustomObject]@{
                添加人 = [string]($worksheet.Cells[$row, $addedByColumn].Value)
                添加时间 = [string]($worksheet.Cells[$row, $addedTimeColumn].Value)
            }
        }
    }
    throw "工作表【$sheet】中不存在规则 $RuleId。"
}


function Get-RuleByIdFromPackage {
    param($Package, [string]$SheetName, [string]$RuleId)
    $sheet = Assert-AllowedSheet -SheetName $SheetName
    if ($RuleId -notmatch '^\d{1,10}$') { throw '规则 ID 无效。' }
    $worksheet = $Package.Workbook.Worksheets[$sheet]
    if ($null -eq $worksheet.Dimension) { throw "工作表【$sheet】中不存在规则 $RuleId。" }
    $headers = Get-ExpectedHeaders -SheetName $sheet
    for ($row = 2; $row -le $worksheet.Dimension.End.Row; $row++) {
        if ([string]($worksheet.Cells[$row, 1].Value) -eq $RuleId) {
            $record = [ordered]@{}
            for ($column = 1; $column -le $headers.Count; $column++) {
                $cell = $worksheet.Cells[$row, $column]
                $record[$headers[$column - 1]] = [string]$cell.Value
                if ($headers[$column - 1] -eq '备注/原因') { $record['备注/原因链接'] = Get-CellHyperlinkUrl -Cell $cell }
            }
            return [PSCustomObject]$record
        }
    }
    throw "工作表【$sheet】中不存在规则 $RuleId。"
}

function Set-RuleMetadataByIdInPackage {
    param($Package, [string]$SheetName, [string]$RuleId, [string]$AddedBy, [string]$AddedTime)
    $sheet = Assert-AllowedSheet -SheetName $SheetName
    if ($RuleId -notmatch '^\d{1,10}$') { throw '规则 ID 无效。' }
    $worksheet = $Package.Workbook.Worksheets[$sheet]
    if ($null -eq $worksheet.Dimension) { throw "工作表【$sheet】中不存在规则 $RuleId。" }
    $headers = Get-ExpectedHeaders -SheetName $sheet
    $idColumn = [array]::IndexOf($headers, 'id') + 1
    $addedByColumn = [array]::IndexOf($headers, '添加人') + 1
    $addedTimeColumn = [array]::IndexOf($headers, '添加时间') + 1
    for ($row = 2; $row -le $worksheet.Dimension.End.Row; $row++) {
        if ([string]($worksheet.Cells[$row, $idColumn].Value) -eq $RuleId) {
            $worksheet.Cells[$row, $addedByColumn].Value = $AddedBy
            $worksheet.Cells[$row, $addedByColumn].Style.Numberformat.Format = '@'
            $worksheet.Cells[$row, $addedTimeColumn].Value = $AddedTime
            $worksheet.Cells[$row, $addedTimeColumn].Style.Numberformat.Format = '@'
            return
        }
    }
    throw "工作表【$sheet】中不存在规则 $RuleId。"
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
    $note = Get-LimitedText -Value $Data.note -FieldName '备注/原因' -MaximumLength 2000
    $addedBy = Get-LimitedText -Value $Data.addedBy -FieldName '添加人' -MaximumLength 128
    $addedTime = Get-LimitedText -Value $Data.addedTime -FieldName '添加时间' -MaximumLength 128
    if (-not $script:TakeoverUnlocked) {
        $addedBy = [string]$env:USERNAME
        if ([string]::IsNullOrWhiteSpace($addedBy)) { $addedBy = [Environment]::UserName }
        $addedBy = Get-LimitedText -Value $addedBy -FieldName '添加人' -MaximumLength 128 -Required
        $addedTime = Get-Date -Format 'yyyy-MM-dd HH:mm'
    } else {
        if ([string]::IsNullOrWhiteSpace($addedBy)) {
            $addedBy = [string]$env:USERNAME
            if ([string]::IsNullOrWhiteSpace($addedBy)) { $addedBy = [Environment]::UserName }
        }
        $addedBy = Get-LimitedText -Value $addedBy -FieldName '添加人' -MaximumLength 128 -Required
        if ([string]::IsNullOrWhiteSpace($addedTime)) { $addedTime = Get-Date -Format 'yyyy-MM-dd HH:mm' }
    }
    $rule = [ordered]@{
        类型 = $itemType
        插件ID = $extensionId
        匹配方式 = $matchType
        软件名关键词 = $namePattern
        版本号 = Get-LimitedText -Value $Data.version -FieldName '版本号' -MaximumLength 2048
        发布者 = Get-LimitedText -Value $Data.publisher -FieldName '发布者' -MaximumLength 256
        '状态/分类' = Get-CanonicalRuleStatus -SheetName $sheet
        '备注/原因' = $note
        '备注/原因链接' = Get-SafeHttpUrl -Value $Data.noteLink
        添加人 = $addedBy
        添加时间 = $addedTime
    }
    return [PSCustomObject]$rule
}


function Get-RuleIdentity {
    param(
        [string]$ItemType,
        [string]$ExtensionId,
        [string]$NamePattern,
        [string]$Publisher = '',
        [string]$Version = ''
    )
    if ($ItemType -eq '软件') {
        $identityRule = [PSCustomObject]@{
            类型 = '软件'
            匹配方式 = '精确'
            软件名关键词 = $NamePattern
            版本号 = $Version
        }
        $effective = Get-EffectiveRuleMatchFields -Rule $identityRule
        $nameKey = (ConvertTo-MatchComparableText $effective.NamePattern).ToLowerInvariant()
        $publisherKey = (ConvertTo-MatchComparableText $Publisher).ToLowerInvariant()
        $versionParts = @((Get-RuleVersionPatterns -RuleVersion ([string]$effective.Version)) | ForEach-Object {
            (ConvertTo-MatchComparableText $_).ToLowerInvariant()
        } | Sort-Object -Unique)
        $versionKey = $versionParts -join ','
        return (ConvertTo-Json -InputObject @('软件', $nameKey, $publisherKey, $versionKey) -Compress)
    }
    return ($ItemType + '|' + ([string]$ExtensionId).Trim().ToLowerInvariant())
}

function Get-WorksheetRowIdentity {
    param($Worksheet, [int]$Row, [string]$SheetName)
    $headers = Get-ExpectedHeaders -SheetName $SheetName
    $map = @{}
    for ($column = 1; $column -le $headers.Count; $column++) { $map[$headers[$column - 1]] = [string]($Worksheet.Cells[$Row, $column].Value) }
    $itemType = if ([string]::IsNullOrWhiteSpace($map['类型'])) { '软件' } else { $map['类型'] }
    return Get-RuleIdentity -ItemType $itemType -ExtensionId $map['插件ID'] -NamePattern $map['软件名关键词'] -Publisher $map['发布者'] -Version $map['版本号']
}

function Remove-IdentitiesFromPackage {
    param($Package, [hashtable]$IdentitySet)
    foreach ($sheetName in $script:AllowedSheets) {
        $worksheet = $Package.Workbook.Worksheets[$sheetName]
        if ($null -eq $worksheet.Dimension) { continue }
        for ($row = $worksheet.Dimension.End.Row; $row -ge 2; $row--) {
            $identity = Get-WorksheetRowIdentity -Worksheet $worksheet -Row $row -SheetName $sheetName
            if ($IdentitySet.ContainsKey($identity)) { $worksheet.DeleteRow($row, 1) }
        }
    }
}

function Get-PackageIdentitySet {
    param($Package)
    $identitySet = @{}
    foreach ($sheetName in $script:AllowedSheets) {
        $worksheet = $Package.Workbook.Worksheets[$sheetName]
        if ($null -eq $worksheet.Dimension) { continue }
        $headers = Get-ExpectedHeaders -SheetName $sheetName
        for ($row = 2; $row -le $worksheet.Dimension.End.Row; $row++) {
            $identity = Get-WorksheetRowIdentity -Worksheet $worksheet -Row $row -SheetName $sheetName
            if (-not [string]::IsNullOrWhiteSpace($identity)) {
                $ruleData = [ordered]@{}
                for ($column = 1; $column -le $headers.Count; $column++) { $ruleData[$headers[$column - 1]] = $worksheet.Cells[$row, $column].Value }
                $identitySet[$identity] = [PSCustomObject]@{ Sheet = $sheetName; Rule = [PSCustomObject]$ruleData }
            }
        }
    }
    return $identitySet
}

function Get-WorksheetAppendState {
    param($Package, [string]$SheetName)
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
    return [PSCustomObject]@{ Sheet = $sheet; Worksheet = $worksheet; Headers = $headers; NextId = $maximumId + 1; NextRow = $targetRow }
}

function Add-RuleToPackageWithState {
    param($State, $Rule)
    if ($State.NextId -gt 9999999999) { throw "工作表【$($State.Sheet)】的规则 ID 已达到上限。" }
    for ($column = 1; $column -le $State.Headers.Count; $column++) {
        $header = $State.Headers[$column - 1]
        if ($header -eq 'id') { $State.Worksheet.Cells[$State.NextRow, $column].Value = [long]$State.NextId }
            else {
            $property = $Rule.PSObject.Properties[$header]
            $value = if ($null -ne $property -and $null -ne $property.Value) { [string]$property.Value } else { '' }
            $State.Worksheet.Cells[$State.NextRow, $column].Value = $value
            if ($header -eq '备注/原因') {
                $noteLink = Get-SafeHttpUrl -Value $Rule.'备注/原因链接'
                if (-not [string]::IsNullOrWhiteSpace($noteLink)) { $State.Worksheet.Cells[$State.NextRow, $column].Hyperlink = [Uri]$noteLink }
            }
            $State.Worksheet.Cells[$State.NextRow, $column].Style.Numberformat.Format = '@'
        }
    }
    $assignedId = [long]$State.NextId
    $State.NextId = [long]$State.NextId + 1
    $State.NextRow = [int]$State.NextRow + 1
    return $assignedId
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
    $state = Get-WorksheetAppendState -Package $Package -SheetName $SheetName
    return Add-RuleToPackageWithState -State $state -Rule $Rule
}

function ConvertTo-MatchComparableText {
    param([object]$Value)
    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return '' }
    $normalized = $text.Normalize([System.Text.NormalizationForm]::FormKC)
    return [regex]::Replace($normalized, '\s+', ' ').Trim()
}

function Get-RuleVersionPatterns {
    param([string]$RuleVersion)
    $text = ConvertTo-MatchComparableText $RuleVersion
    if ([string]::IsNullOrWhiteSpace($text)) { return @() }
    if (($text.StartsWith('[') -and $text.EndsWith(']')) -or ($text.StartsWith('(') -and $text.EndsWith(')'))) {
        $text = $text.Substring(1, $text.Length - 2)
    }
    $patterns = @()
    foreach ($part in [regex]::Split($text, '[,，;；\r\n]+')) {
        $pattern = (ConvertTo-MatchComparableText $part).Trim('"', "'", ' ')
        if (-not [string]::IsNullOrWhiteSpace($pattern)) { $patterns += $pattern }
    }
    if ($patterns.Count -eq 0) { return @($text) }
    return $patterns
}

function Test-VersionMatch {
    param([object]$InstalledVersion, [string]$RuleVersion)
    if ([string]::IsNullOrWhiteSpace($RuleVersion)) { return $true }
    $rulePatterns = @(Get-RuleVersionPatterns -RuleVersion $RuleVersion)
    foreach ($candidate in @($InstalledVersion)) {
        $candidateText = ConvertTo-MatchComparableText $candidate
        if ([string]::IsNullOrWhiteSpace($candidateText)) { continue }
        foreach ($rulePattern in $rulePatterns) {
            if ($candidateText -like $rulePattern) { return $true }
        }
    }
    return $false
}

function Get-EffectiveRuleMatchFields {
    param($Rule)
    $matchType = [string]$Rule.'匹配方式'
    $namePattern = [string]$Rule.'软件名关键词'
    $version = [string]$Rule.'版本号'
    $embeddedVersion = $false

    if (([string]$Rule.'类型' -eq '软件' -or [string]::IsNullOrWhiteSpace([string]$Rule.'类型')) -and
        -not [string]::IsNullOrWhiteSpace($namePattern)) {
        $nameVersionMatch = [regex]::Match(
            $namePattern.Trim(),
            '^(?<base>.+?)(?:\s+|\s*[-–—_/]\s*)(?:v(?:ersion)?\s*)?\(?(?<version>\d+(?:\.\d+)+(?:[-+][0-9A-Za-z.-]+)?)\)?$',
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
        )
        if ($nameVersionMatch.Success -and -not [string]::IsNullOrWhiteSpace($nameVersionMatch.Groups['base'].Value)) {
            $namePattern = $nameVersionMatch.Groups['base'].Value.Trim()
            if ([string]::IsNullOrWhiteSpace($version)) { $version = $nameVersionMatch.Groups['version'].Value }
            $embeddedVersion = $true
        }
    }

    return [PSCustomObject]@{
        MatchType = $matchType
        NamePattern = $namePattern
        Version = $version
        EmbeddedVersion = $embeddedVersion
    }
}

function Test-NameMatch {
    param([string]$InstalledName, $Rule)
    if ($Rule.'匹配方式' -eq '插件ID精确') {
        if ([string]::IsNullOrWhiteSpace([string]$Rule.'插件ID')) { return $false }
        return [string]::Equals($InstalledName, [string]$Rule.'插件ID', [System.StringComparison]::OrdinalIgnoreCase)
    }
    $effectiveRule = Get-EffectiveRuleMatchFields -Rule $Rule
    $pattern = [string]$effectiveRule.NamePattern
    if ([string]::IsNullOrWhiteSpace($pattern)) { return $false }
    $installedComparable = ConvertTo-MatchComparableText $InstalledName
    $patternComparable = ConvertTo-MatchComparableText $pattern
    switch ([string]$effectiveRule.MatchType) {
        '精确' {
            if ($effectiveRule.EmbeddedVersion) {
                $installedBaseName = [regex]::Replace(
                    $installedComparable,
                    '(?i)(?:\s+|\s*[-–—_/]\s*)(?:v(?:ersion)?\s*)?\(?\d+(?:\.\d+)+(?:[-+][0-9A-Za-z.-]+)?\)?$',
                    ''
                ).Trim()
                return [string]::Equals($installedBaseName, $patternComparable, [System.StringComparison]::OrdinalIgnoreCase)
            }
            return [string]::Equals($installedComparable, $patternComparable, [System.StringComparison]::OrdinalIgnoreCase)
        }
        '包含' { return $installedComparable.IndexOf($patternComparable, [System.StringComparison]::OrdinalIgnoreCase) -ge 0 }
        '通配符' { return $installedComparable -like $patternComparable }
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
    $installedComparable = ConvertTo-MatchComparableText $InstalledPublisher
    $publisherComparable = ConvertTo-MatchComparableText $publisher
    return $installedComparable.IndexOf($publisherComparable, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
}

function New-RuleMatcher {
    param($Rules, [string]$ItemType)
    $exactNames = @{}
    $extensionIds = @{}
    $fallback = @()
    $order = 0
    foreach ($rule in @($Rules)) {
        $entry = [PSCustomObject]@{ Order = $order; Rule = $rule }
        $order++
        $ruleType = if ([string]::IsNullOrWhiteSpace([string]$rule.'类型')) { '软件' } else { [string]$rule.'类型' }
        if ($ruleType -ne $ItemType) { continue }
        $effectiveRule = Get-EffectiveRuleMatchFields -Rule $rule
        $matchType = [string]$effectiveRule.MatchType
        if ($matchType -eq '插件ID精确') {
            $key = ([string]$rule.'插件ID').Trim()
            if (-not $extensionIds.ContainsKey($key)) { $extensionIds[$key] = @() }
            $extensionIds[$key] = @($extensionIds[$key]) + $entry
        } elseif ($matchType -eq '精确' -and -not $effectiveRule.EmbeddedVersion) {
            $key = ConvertTo-MatchComparableText $effectiveRule.NamePattern
            if (-not $exactNames.ContainsKey($key)) { $exactNames[$key] = @() }
            $exactNames[$key] = @($exactNames[$key]) + $entry
        } else {
            $fallback += $entry
        }
    }
    return [PSCustomObject]@{ IsRuleMatcher = $true; ExactNames = $exactNames; ExtensionIds = $extensionIds; Fallback = $fallback }
}

function Find-FirstMatchingRule {
    param($Rules, [string]$ItemType, [string]$DisplayName, [string]$Publisher, [string]$ExtensionId, [object]$Version, [switch]$IgnoreVersion, [switch]$IgnorePublisher)
    $candidateRules = @($Rules)
    if ($null -ne $Rules -and $Rules.IsRuleMatcher -eq $true) {
        $candidateEntries = @()
        if (-not [string]::IsNullOrWhiteSpace($ExtensionId) -and $Rules.ExtensionIds.ContainsKey($ExtensionId)) { $candidateEntries += @($Rules.ExtensionIds[$ExtensionId]) }
        $normalizedDisplayName = ConvertTo-MatchComparableText $DisplayName
        if ($Rules.ExactNames.ContainsKey($normalizedDisplayName)) { $candidateEntries += @($Rules.ExactNames[$normalizedDisplayName]) }
        $candidateEntries += @($Rules.Fallback)
        $candidateRules = @($candidateEntries | Sort-Object Order | ForEach-Object { $_.Rule })
    }
    foreach ($rule in $candidateRules) {
        $ruleType = if ([string]::IsNullOrWhiteSpace([string]$rule.'类型')) { '软件' } else { [string]$rule.'类型' }
        if ($ruleType -ne $ItemType) { continue }
        $testName = if ($rule.'匹配方式' -eq '插件ID精确') { $ExtensionId } else { $DisplayName }
        if (-not (Test-NameMatch -InstalledName $testName -Rule $rule)) { continue }
        if (-not $IgnorePublisher -and -not (Test-PublisherMatch -InstalledPublisher $Publisher -RulePublisher $rule.'发布者')) { continue }
        $effectiveRule = Get-EffectiveRuleMatchFields -Rule $rule
        if (-not $IgnoreVersion -and -not (Test-VersionMatch -InstalledVersion $Version -RuleVersion ([string]$effectiveRule.Version))) { continue }
        return $rule
    }
    return $null
}

function Find-BlacklistedPluginNameCollision {
    param($Rules, [string]$ItemType, [string]$DisplayName, [string]$ExtensionId)
    if ($ItemType -eq '软件' -or [string]::IsNullOrWhiteSpace($DisplayName)) { return $null }
    $installedName = ConvertTo-MatchComparableText $DisplayName
    foreach ($rule in @($Rules)) {
        $ruleType = if ([string]::IsNullOrWhiteSpace([string]$rule.'类型')) { '软件' } else { [string]$rule.'类型' }
        if ($ruleType -ne $ItemType -or [string]$rule.'匹配方式' -ne '插件ID精确') { continue }
        $ruleExtensionId = ([string]$rule.'插件ID').Trim()
        if ([string]::Equals($ruleExtensionId, ([string]$ExtensionId).Trim(), [System.StringComparison]::OrdinalIgnoreCase)) { continue }
        $ruleName = ConvertTo-MatchComparableText ([string]$rule.'软件名关键词')
        if ([string]::IsNullOrWhiteSpace($ruleName)) { continue }
        if ([string]::Equals($ruleName, $installedName, [System.StringComparison]::OrdinalIgnoreCase)) { return $rule }
    }
    return $null
}

function New-ComplianceItem {
    param($Source, [string]$ItemType, [string]$Status, [string]$Reason, $MatchedRule, [string]$MatchedSheet)
    $isSoftware = $ItemType -eq '软件'
    $iconPath = if ($isSoftware) { [string]$Source.图标路径 } else { [string]$Source.IconPath }
    $iconIndex = if ($isSoftware) { [int]$Source.图标索引 } else { 0 }
    $displayName = if ($isSoftware) { [string]$Source.名称 } else { [string]$Source.Name }
    $extensionId = if ($isSoftware) { '' } else { [string]$Source.ExtensionId }
    $iconIdentity = '{0}|{1}|{2}' -f $ItemType, $displayName.ToLowerInvariant(), $extensionId.ToLowerInvariant()
    if ($script:IconKeyByIdentity.ContainsKey($iconIdentity)) {
        $lazyIconKey = [string]$script:IconKeyByIdentity[$iconIdentity]
    } else {
        $lazyIconKey = New-RandomToken -ByteCount 12
        $script:IconKeyByIdentity[$iconIdentity] = $lazyIconKey
        $script:IconSourceByKey[$lazyIconKey] = [PSCustomObject]@{ ItemType = $ItemType; Source = $Source; Path = $iconPath; Index = $iconIndex; DataUri = '' }
    }
    $script:ReportIconData[$iconIdentity] = [PSCustomObject]@{ ItemType = $ItemType; Name = $displayName; Publisher = if ($isSoftware) { [string]$Source.发布者 } else { [string]$Source.Publisher }; ExtensionId = $extensionId; IconKey = $lazyIconKey }
    if ($null -ne $MatchedRule) {
        $ruleType = if ([string]::IsNullOrWhiteSpace([string]$MatchedRule.'类型')) { $ItemType } else { [string]$MatchedRule.'类型' }
        $ruleIdentity = Get-RuleIdentity -ItemType $ruleType -ExtensionId ([string]$MatchedRule.'插件ID') -NamePattern ([string]$MatchedRule.'软件名关键词') -Publisher ([string]$MatchedRule.'发布者') -Version ([string]$MatchedRule.'版本号')
        $script:RuleIconCache[$ruleIdentity] = [PSCustomObject]@{ Path = $iconPath; Index = $iconIndex; DataUri = ''; IconKey = $lazyIconKey }
        if (-not [string]::IsNullOrWhiteSpace([string]$MatchedRule.id)) {
            $script:RuleIconDataById[[string]$MatchedRule.id] = $lazyIconKey
        }
    }
    return [PSCustomObject]@{
        名称 = if ($isSoftware) { [string]$Source.名称 } else { [string]$Source.Name }
        版本 = if ($isSoftware) { [string]$Source.版本 } else { [string]$Source.Version }
        发布者 = if ($isSoftware) { [string]$Source.发布者 } else { [string]$Source.Publisher }
        安装日期 = if ($isSoftware) { [string]$Source.安装日期 } else { '' }
        安装路径 = if ($isSoftware) { [string]$Source.安装路径 } else { '' }
        状态 = $Status
        原因 = $Reason
        '备注/原因链接' = if ($null -ne $MatchedRule) { [string]$MatchedRule.'备注/原因链接' } else { '' }
        类型 = $ItemType
        插件ID = if ($isSoftware) { '' } else { [string]$Source.ExtensionId }
        Locations = if ($isSoftware) { $null } else { $Source.Locations }
        图标路径 = $iconPath
        图标索引 = $iconIndex
        图标数据 = ''
        图标键 = $lazyIconKey
        MatchedSheet = $MatchedSheet
        MatchedRule = $MatchedRule
        MatchedRuleId = if ($null -ne $MatchedRule) { $MatchedRule.id } else { $null }
    }
}

function Get-ComplianceResult {
    param($Installed, $BlackRules, $WhiteRules, $PendingRules, [string]$ItemType)
    $results = @()
    $blackMatcher = New-RuleMatcher -Rules $BlackRules -ItemType $ItemType
    $whiteMatcher = New-RuleMatcher -Rules $WhiteRules -ItemType $ItemType
    $pendingMatcher = New-RuleMatcher -Rules $PendingRules -ItemType $ItemType
    foreach ($item in @($Installed)) {
        $displayName = if ($ItemType -eq '软件') { [string]$item.名称 } else { [string]$item.Name }
        $publisher = if ($ItemType -eq '软件') { [string]$item.发布者 } else { [string]$item.Publisher }
        $version = if ($ItemType -eq '软件') { [string]$item.版本 } elseif ($null -ne $item.Versions) { @($item.Versions) } else { [string]$item.Version }
        $extensionId = if ($ItemType -eq '软件') { '' } else { [string]$item.ExtensionId }

        # 黑名单不因版本升级自动失效；软件仍遵守发布者限制，插件只依赖其规则匹配方式（ID 规则必须精确匹配 ID）。
        $black = if ($ItemType -eq '软件') {
            Find-FirstMatchingRule -Rules $blackMatcher -ItemType $ItemType -DisplayName $displayName -Publisher $publisher -ExtensionId $extensionId -Version $version -IgnoreVersion
        } else {
            Find-FirstMatchingRule -Rules $blackMatcher -ItemType $ItemType -DisplayName $displayName -Publisher $publisher -ExtensionId $extensionId -Version $version -IgnoreVersion -IgnorePublisher
        }
        if ($null -ne $black) {
            $reason = Get-UserNoteText $black.'备注/原因'
            $results += New-ComplianceItem -Source $item -ItemType $ItemType -Status '命中黑名单' -Reason $reason -MatchedRule $black -MatchedSheet '黑名单'

            continue
        }

        $white = Find-FirstMatchingRule -Rules $whiteMatcher -ItemType $ItemType -DisplayName $displayName -Publisher $publisher -ExtensionId $extensionId -Version $version
        if ($null -ne $white) {
            $reason = Get-UserNoteText $white.'备注/原因'
            $results += New-ComplianceItem -Source $item -ItemType $ItemType -Status '已匹配' -Reason $reason -MatchedRule $white -MatchedSheet '白名单'
            continue
        }
        # 如果找不到精确匹配版本的白名单，尝试寻找名称匹配但版本不限/不匹配的规则，标记为“版本变化”
        $whiteAnyVersion = Find-FirstMatchingRule -Rules $whiteMatcher -ItemType $ItemType -DisplayName $displayName -Publisher $publisher -ExtensionId $extensionId -Version $version -IgnoreVersion
        if ($null -ne $whiteAnyVersion) {
            $results += New-ComplianceItem -Source $item -ItemType $ItemType -Status '版本变化' -Reason '' -MatchedRule $whiteAnyVersion -MatchedSheet '白名单'
            continue
        }

        $whitePublisherMismatch = Find-FirstMatchingRule -Rules $whiteMatcher -ItemType $ItemType -DisplayName $displayName -Publisher $publisher -ExtensionId $extensionId -Version $version -IgnorePublisher
        if ($null -ne $whitePublisherMismatch -and -not [string]::IsNullOrWhiteSpace([string]$whitePublisherMismatch.'发布者')) {
            $expectedPublisher = [string]$whitePublisherMismatch.'发布者'
            $actualPublisher = if ([string]::IsNullOrWhiteSpace($publisher)) { '空' } else { $publisher }
            $reason = "白名单名称和版本一致，但发布者不同。规则发布者：$expectedPublisher；扫描发布者：$actualPublisher"
            $results += New-ComplianceItem -Source $item -ItemType $ItemType -Status '待定' -Reason $reason -MatchedRule $whitePublisherMismatch -MatchedSheet '白名单'
            continue
        }

        $whiteMultipleMismatch = Find-FirstMatchingRule -Rules $whiteMatcher -ItemType $ItemType -DisplayName $displayName -Publisher $publisher -ExtensionId $extensionId -Version $version -IgnorePublisher -IgnoreVersion
        if ($null -ne $whiteMultipleMismatch -and -not [string]::IsNullOrWhiteSpace([string]$whiteMultipleMismatch.'发布者')) {
            $effectiveMultipleRule = Get-EffectiveRuleMatchFields -Rule $whiteMultipleMismatch
            $actualPublisher = if ([string]::IsNullOrWhiteSpace($publisher)) { '空' } else { $publisher }
            $reason = "白名单名称一致，但版本和发布者均不同。规则版本：$($effectiveMultipleRule.Version)；扫描版本：$version；规则发布者：$($whiteMultipleMismatch.'发布者')；扫描发布者：$actualPublisher"
            $results += New-ComplianceItem -Source $item -ItemType $ItemType -Status '待定' -Reason $reason -MatchedRule $whiteMultipleMismatch -MatchedSheet '白名单'
            continue
        }

        if ($ItemType -ne '软件') {
            $blackNameCollision = Find-BlacklistedPluginNameCollision -Rules $BlackRules -ItemType $ItemType -DisplayName $displayName -ExtensionId $extensionId
            if ($null -ne $blackNameCollision) {
                $blackId = [string]$blackNameCollision.'插件ID'
                $reason = "名称与黑名单插件相同，但插件 ID 不同，请人工确认。黑名单 ID：$blackId；当前 ID：$extensionId"
                $results += New-ComplianceItem -Source $item -ItemType $ItemType -Status '待定' -Reason $reason -MatchedRule $null -MatchedSheet ''
                continue
            }
        }

        $pending = Find-FirstMatchingRule -Rules $pendingMatcher -ItemType $ItemType -DisplayName $displayName -Publisher $publisher -ExtensionId $extensionId -Version $version
        if ($null -ne $pending) {
            $reason = Get-UserNoteText $pending.'备注/原因'
            $results += New-ComplianceItem -Source $item -ItemType $ItemType -Status '待定' -Reason $reason -MatchedRule $pending -MatchedSheet '待定'

            continue
        }

                $results += New-ComplianceItem -Source $item -ItemType $ItemType -Status '待定' -Reason '' -MatchedRule $null -MatchedSheet ''

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
        $existingIdentities = Get-PackageIdentitySet -Package $package
        $pendingState = Get-WorksheetAppendState -Package $package -SheetName '待定'
        foreach ($item in $candidates) {
            $itemType = [string]$item.类型
            $identity = Get-RuleIdentity -ItemType $itemType -ExtensionId ([string]$item.插件ID) -NamePattern ([string]$item.名称) -Publisher ([string]$item.发布者) -Version ([string]$item.版本)
            if ($existingIdentities.ContainsKey($identity)) {
                $existingRule = $existingIdentities[$identity]
                $item.MatchedSheet = $existingRule.Sheet
                $item.MatchedRuleId = $existingRule.Rule.id
                $item.MatchedRule = $existingRule.Rule
                continue
            }
            $ruleData = [PSCustomObject]@{
                type = $itemType
                extId = [string]$item.插件ID
                matchType = if ($itemType -eq '软件') { '精确' } else { '插件ID精确' }
                namePattern = [string]$item.名称
                version = [string]$item.版本
                                publisher = [string]$item.发布者
                note = ''

            }
            $rule = Convert-InputToRule -Data $ruleData -TargetSheet '待定'
            $newId = Add-RuleToPackageWithState -State $pendingState -Rule $rule
            $rule | Add-Member -NotePropertyName "id" -NotePropertyValue $newId -Force
            $item.MatchedRule = $rule
            $item.MatchedRuleId = [string]$newId
            $item.MatchedSheet = "待定"
            $existingIdentities[$identity] = [PSCustomObject]@{ Sheet = '待定'; Rule = $rule }
            $added++
        }
        return $added
    }
    $count = Invoke-WorkbookTransaction -Path $Path -Operation $operation
    if ($count -gt 0) { Remove-UndoSnapshot }
    return $count
}

function Clear-SystemGeneratedRuleNotes {
    param([string]$Path)
    $probePackage = $null
    $needsClear = $false
    try {
        $probePackage = Open-ExcelPackage -Path $Path -ErrorAction Stop
        Assert-WorkbookPackageSchema -Package $probePackage
        foreach ($sheetName in $script:AllowedSheets) {
            $worksheet = $probePackage.Workbook.Worksheets[$sheetName]
            if ($null -eq $worksheet.Dimension) { continue }
            $headers = Get-ExpectedHeaders -SheetName $sheetName
            $noteColumn = [array]::IndexOf($headers, '备注/原因') + 1
            for ($row = 2; $row -le $worksheet.Dimension.End.Row; $row++) {
                $rawNote = [string]($worksheet.Cells[$row, $noteColumn].Value)
                if ($rawNote -ne (Get-UserNoteText $rawNote)) {
                    $needsClear = $true
                    break
                }
            }
            if ($needsClear) { break }
        }
    } finally {
        if ($null -ne $probePackage) { $probePackage.Dispose() }
    }
    if (-not $needsClear) { return 0 }
    $operation = {
        param($package)
        $cleared = 0
        foreach ($sheetName in $script:AllowedSheets) {
            $worksheet = $package.Workbook.Worksheets[$sheetName]
            if ($null -eq $worksheet.Dimension) { continue }
            $headers = Get-ExpectedHeaders -SheetName $sheetName
            $noteColumn = [array]::IndexOf($headers, '备注/原因') + 1
            for ($row = 2; $row -le $worksheet.Dimension.End.Row; $row++) {
                $cell = $worksheet.Cells[$row, $noteColumn]
                if (([string]$cell.Value) -ne (Get-UserNoteText $cell.Value)) {
                    $cell.Value = ''
                    $cell.Hyperlink = $null
                    $cell.Style.Numberformat.Format = '@'
                    $cleared++
                }
            }
        }
        return $cleared
    }
    return Invoke-WorkbookTransaction -Path $Path -Operation $operation
}

function Get-InventoryData {
    param([switch]$ForceRefresh, [switch]$IncludeSystem)
    if ($ForceRefresh) {
        Start-InventoryScanJobs -IncludeSystem:$IncludeSystem -IncludeInactive:$IncludeInactiveExtensions -AllUsers:$ScanAllUsers
    }
    if (-not $ForceRefresh -and $null -ne $script:InventoryCache) {
        return $script:InventoryCache
    }
    if ($script:InventoryScanJobs.Count -eq 0) {
        Start-InventoryScanJobs -IncludeSystem:$IncludeSystem -IncludeInactive:$IncludeInactiveExtensions -AllUsers:$ScanAllUsers
    }
    while ($null -eq $script:InventoryCache -and [string]::IsNullOrWhiteSpace($script:InventoryScanError)) {
        $null = Update-InventoryScanState
        if ($null -eq $script:InventoryCache) { Start-Sleep -Milliseconds 100 }
    }
    if (-not [string]::IsNullOrWhiteSpace($script:InventoryScanError)) { throw $script:InventoryScanError }
    return $script:InventoryCache
}

function Stop-InventoryScanJobs {
    foreach ($entry in @($script:InventoryScanJobs.Values)) {
        $job = $entry.Job
        if ($null -eq $job) { continue }
        try { if ($job.State -in @('Running', 'NotStarted')) { Stop-Job -Job $job -ErrorAction SilentlyContinue } } catch {}
        try { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue } catch {}
    }
    $script:InventoryScanJobs = @{}
}

function Start-InventoryScanJobs {
    param([switch]$IncludeSystem, [switch]$IncludeInactive, [switch]$AllUsers)
    Stop-InventoryScanJobs
    $script:InventoryCache = $null
    $script:InventoryScanResults = @{}
    $script:InventoryScanError = ''
    $script:InventoryScanStartedAt = Get-Date
    foreach ($cache in @($script:RuleIconCache, $script:RuleIconDataById, $script:ReportIconData, $script:IconSourceByKey, $script:IconKeyByIdentity)) {
        if ($null -ne $cache) { $cache.Clear() }
    }
    $rootPath = $PSScriptRoot
    $specifications = @(
        [PSCustomObject]@{ Kind = 'Software'; Script = {
            param($RootPath, $IncludeSystemValue, $IncludeInactiveValue, $AllUsersValue)
            . (Join-Path $RootPath 'Get-InstalledSoftware.ps1')
            [PSCustomObject]@{ Kind = 'Software'; Items = @(Get-InstalledSoftwareList -IncludeSystemComponents:$IncludeSystemValue -AllUsers:$AllUsersValue) }
        } },
        [PSCustomObject]@{ Kind = 'Chromium'; Script = {
            param($RootPath, $IncludeSystemValue, $IncludeInactiveValue, $AllUsersValue)
            . (Join-Path $RootPath 'Get-InstalledExtensions.ps1')
            [PSCustomObject]@{ Kind = 'Chromium'; Items = @(Get-ChromiumExtensions -IncludeInactive:$IncludeInactiveValue -AllUsers:$AllUsersValue) }
        } },
        [PSCustomObject]@{ Kind = 'Firefox'; Script = {
            param($RootPath, $IncludeSystemValue, $IncludeInactiveValue, $AllUsersValue)
            . (Join-Path $RootPath 'Get-InstalledExtensions.ps1')
            [PSCustomObject]@{ Kind = 'Firefox'; Items = @(Get-FirefoxExtensions -IncludeInactive:$IncludeInactiveValue -AllUsers:$AllUsersValue) }
        } }
    )
    foreach ($specification in $specifications) {
        try {
            $job = Start-Job -ScriptBlock $specification.Script -ArgumentList @($rootPath, $IncludeSystem.IsPresent, $IncludeInactive.IsPresent, $AllUsers.IsPresent) -ErrorAction Stop
            $script:InventoryScanJobs[$specification.Kind] = [PSCustomObject]@{ Job = $job; Received = $false }
        } catch {
            Stop-InventoryScanJobs
            $script:InventoryScanError = "无法启动后台扫描：$($_.Exception.Message)"
            break
        }
    }
}

function Update-InventoryScanState {
    if ($null -ne $script:InventoryCache) { return $true }
    foreach ($kind in @($script:InventoryScanJobs.Keys)) {
        $entry = $script:InventoryScanJobs[$kind]
        if ($entry.Received) { continue }
        $job = $entry.Job
        if ($job.State -eq 'Failed') {
            $reason = if ($null -ne $job.ChildJobs[0].JobStateInfo.Reason) { $job.ChildJobs[0].JobStateInfo.Reason.Message } else { '未知错误' }
            $script:InventoryScanError = "$kind 扫描失败：$reason"
            return $false
        }
        if ($job.State -eq 'Completed') {
            try {
                $payload = @(Receive-Job -Job $job -ErrorAction Stop) | Select-Object -Last 1
                if ($null -eq $payload -or [string]$payload.Kind -ne $kind) { throw '后台扫描没有返回有效结果。' }
                $script:InventoryScanResults[$kind] = @($payload.Items)
                $entry.Received = $true
            } catch {
                $script:InventoryScanError = "$kind 扫描结果读取失败：$($_.Exception.Message)"
                return $false
            }
        }
    }
    if ($script:InventoryScanJobs.Count -eq 3 -and @($script:InventoryScanJobs.Values | Where-Object { -not $_.Received }).Count -eq 0) {
        $script:InventoryCache = [PSCustomObject]@{
            Software = @($script:InventoryScanResults.Software)
            Chromium = @($script:InventoryScanResults.Chromium)
            Firefox = @($script:InventoryScanResults.Firefox)
            CreatedAt = Get-Date
        }
        Stop-InventoryScanJobs
        return $true
    }
    return $false
}

function Get-InventoryScanStatus {
    $ready = Update-InventoryScanState
    $completed = @($script:InventoryScanJobs.Values | Where-Object { $_.Received }).Count
    if ($ready) { $completed = 3 }
    $elapsed = if ($null -ne $script:InventoryScanStartedAt) { [Math]::Round(((Get-Date) - $script:InventoryScanStartedAt).TotalSeconds, 1) } else { 0 }
    return [PSCustomObject]@{ Ready = $ready; Completed = $completed; Total = 3; Error = $script:InventoryScanError; ElapsedSeconds = $elapsed }
}

function Initialize-NativeIconApi {
    if ($null -ne ('CheckSentry.NativeIconMethods' -as [type])) { return }
    $source = @'
using System;
using System.Runtime.InteropServices;

namespace CheckSentry {
    public static class NativeIconMethods {
        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        public static extern uint ExtractIconEx(
            string szFileName,
            int nIconIndex,
            out IntPtr phiconLarge,
            out IntPtr phiconSmall,
            uint nIcons);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool DestroyIcon(IntPtr hIcon);
    }
}
'@
    Add-Type -TypeDefinition $source -ErrorAction Stop
}

function Get-LocalIconDataUri {
    param(
        [string]$Path,
        [int]$IconIndex = 0
    )
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) { return '' }
    try {
        $fileItem = Get-Item -LiteralPath $Path -ErrorAction Stop
        $cacheKey = '{0}|{1}|{2}|{3}' -f $fileItem.FullName.ToLowerInvariant(), $IconIndex, $fileItem.Length, $fileItem.LastWriteTimeUtc.Ticks
        if ($script:IconDataCache.ContainsKey($cacheKey)) { return [string]$script:IconDataCache[$cacheKey] }
        if ($script:IconDataCache.Count -gt 2048) { $script:IconDataCache.Clear() }

        $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
        if ($extension -in @('.exe', '.dll')) {
            Add-Type -AssemblyName System.Drawing -ErrorAction Stop
            Initialize-NativeIconApi
            $largeHandle = [IntPtr]::Zero
            $smallHandle = [IntPtr]::Zero
            $icon = $null; $bitmap = $null; $stream = $null
            try {
                $count = [CheckSentry.NativeIconMethods]::ExtractIconEx($Path, $IconIndex, [ref]$largeHandle, [ref]$smallHandle, 1)
                $selectedHandle = if ($largeHandle -ne [IntPtr]::Zero) { $largeHandle } else { $smallHandle }
                if ($count -gt 0 -and $selectedHandle -ne [IntPtr]::Zero) {
                    $icon = [System.Drawing.Icon]::FromHandle($selectedHandle)
                } else {
                    $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($Path)
                }
                if ($null -eq $icon) { $script:IconDataCache[$cacheKey] = ''; return '' }
                $bitmap = $icon.ToBitmap()
                $stream = New-Object System.IO.MemoryStream
                $bitmap.Save($stream, [System.Drawing.Imaging.ImageFormat]::Png)
                $dataUri = 'data:image/png;base64,' + [Convert]::ToBase64String($stream.ToArray())
                $script:IconDataCache[$cacheKey] = $dataUri
                return $dataUri
            } finally {
                if ($null -ne $stream) { $stream.Dispose() }
                if ($null -ne $bitmap) { $bitmap.Dispose() }
                if ($null -ne $icon) { $icon.Dispose() }
                if ($largeHandle -ne [IntPtr]::Zero) { $null = [CheckSentry.NativeIconMethods]::DestroyIcon($largeHandle) }
                if ($smallHandle -ne [IntPtr]::Zero -and $smallHandle -ne $largeHandle) { $null = [CheckSentry.NativeIconMethods]::DestroyIcon($smallHandle) }
            }
        }
        $bytes = $null
        if ($extension -notin @('.png', '.jpg', '.jpeg', '.gif', '.ico')) {
            if ([string]::IsNullOrWhiteSpace($extension)) {
                $bytes = [System.IO.File]::ReadAllBytes($Path)
                if ($bytes.Length -ge 4 -and $bytes[0] -eq 0 -and $bytes[1] -eq 0 -and $bytes[2] -eq 1 -and $bytes[3] -eq 0) {
                    $extension = '.ico'
                }
            }
            if ($extension -notin @('.png', '.jpg', '.jpeg', '.gif', '.ico')) { $script:IconDataCache[$cacheKey] = ''; return '' }
        }
        if ($null -eq $bytes) { $bytes = [System.IO.File]::ReadAllBytes($Path) }
        if ($bytes.Length -eq 0 -or $bytes.Length -gt 524288) { $script:IconDataCache[$cacheKey] = ''; return '' }
        $mime = ''
        if ($extension -eq '.png' -and $bytes.Length -ge 8 -and $bytes[0] -eq 0x89 -and $bytes[1] -eq 0x50 -and $bytes[2] -eq 0x4E -and $bytes[3] -eq 0x47) { $mime = 'image/png' }
        elseif ($extension -in @('.jpg', '.jpeg') -and $bytes.Length -ge 3 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xD8 -and $bytes[2] -eq 0xFF) { $mime = 'image/jpeg' }
        elseif ($extension -eq '.gif' -and $bytes.Length -ge 6 -and [Text.Encoding]::ASCII.GetString($bytes, 0, 6) -match '^GIF8[79]a$') { $mime = 'image/gif' }
        elseif ($extension -eq '.ico' -and $bytes.Length -ge 4 -and $bytes[0] -eq 0 -and $bytes[1] -eq 0 -and $bytes[2] -eq 1 -and $bytes[3] -eq 0) { $mime = 'image/x-icon' }
        if ([string]::IsNullOrWhiteSpace($mime)) { $script:IconDataCache[$cacheKey] = ''; return '' }
        $dataUri = 'data:' + $mime + ';base64,' + [Convert]::ToBase64String($bytes)
        $script:IconDataCache[$cacheKey] = $dataUri
        return $dataUri
    } catch {
        Write-Verbose "图标提取失败：$Path，索引 $IconIndex。$($_.Exception.Message)"
        return ''
    }
}

function Get-RuleIconDataUri {
    param($Rule, [bool]$IncludeSystem)
    $key = Get-RuleIconKey -Rule $Rule
    if ([string]::IsNullOrWhiteSpace($key)) { return '' }
    return Get-IconDataUriByKey -Key $key
}

function Prime-RuleIconCacheFromReport {
    param([object]$Inventory, [object]$AllRules, [bool]$IncludeSystem)
    $black = @($AllRules.'黑名单')
    $white = @($AllRules.'白名单')
    $pending = @($AllRules.'待定')
    $null = @(Get-ComplianceResult -Installed $Inventory.Software -BlackRules $black -WhiteRules $white -PendingRules $pending -ItemType '软件')
    $null = @(Get-ComplianceResult -Installed $Inventory.Chromium -BlackRules $black -WhiteRules $white -PendingRules $pending -ItemType 'Chromium插件')
    $null = @(Get-ComplianceResult -Installed $Inventory.Firefox -BlackRules $black -WhiteRules $white -PendingRules $pending -ItemType 'Firefox插件')
}

function Get-RuleIconKey {
    param($Rule)
    $ruleId = [string]$Rule.id
    if (-not [string]::IsNullOrWhiteSpace($ruleId) -and $script:RuleIconDataById.ContainsKey($ruleId)) {
        return [string]$script:RuleIconDataById[$ruleId]
    }
    $itemType = if ([string]::IsNullOrWhiteSpace([string]$Rule.'类型')) { '软件' } else { [string]$Rule.'类型' }
    foreach ($reportRecord in @($script:ReportIconData.Values)) {
        if ([string]$reportRecord.ItemType -ne $itemType) { continue }
        if ($itemType -eq '软件') {
            if (-not (Test-NameMatch -InstalledName ([string]$reportRecord.Name) -Rule $Rule)) { continue }
            if (-not (Test-PublisherMatch -InstalledPublisher ([string]$reportRecord.Publisher) -RulePublisher $Rule.'发布者')) { continue }
            return [string]$reportRecord.IconKey
        }
        if ([string]::Equals([string]$reportRecord.ExtensionId, [string]$Rule.'插件ID', [System.StringComparison]::OrdinalIgnoreCase)) {
            return [string]$reportRecord.IconKey
        }
    }
    return ''
}

function Get-IconDataUriByKey {
    param([string]$Key)
    if ($Key -notmatch '^[a-f0-9]{24}$' -or -not $script:IconSourceByKey.ContainsKey($Key)) { return '' }
    $record = $script:IconSourceByKey[$Key]
    if (-not [string]::IsNullOrWhiteSpace([string]$record.DataUri)) { return [string]$record.DataUri }
    $path = [string]$record.Path
    $index = [int]$record.Index
    if ([string]$record.ItemType -eq '软件' -and ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path -LiteralPath $path -PathType Leaf))) {
        $reference = Resolve-InstalledSoftwareIconReference -Item $record.Source
        if ($null -ne $reference) {
            $path = [string]$reference.Path
            $index = [int]$reference.Index
            $record.Path = $path
            $record.Index = $index
        }
    }
    $record.DataUri = Get-LocalIconDataUri -Path $path -IconIndex $index
    return [string]$record.DataUri
}

function Convert-RulesForWeb {
    param([object[]]$Rules, [bool]$IncludeSystem)

    $output = @()
    foreach ($rule in @($Rules)) {
        $ordered = [ordered]@{}
        foreach ($property in $rule.PSObject.Properties) { $ordered[$property.Name] = $property.Value }
        $ordered['图标数据'] = ''
        $ordered['图标键'] = Get-RuleIconKey -Rule $rule
        $output += [PSCustomObject]$ordered
    }
    return $output
}

function ConvertTo-HtmlEncodedText {
    param([object]$Value)
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function ConvertTo-SafeNoteHtml {
    param([object]$Text, [object]$Link)
    $textValue = if ($null -eq $Text) { '' } else { [string]$Text }
    $encodedText = ConvertTo-HtmlEncodedText $textValue
    $safeLink = Get-SafeHttpUrl -Value $Link
    if (-not [string]::IsNullOrWhiteSpace($safeLink) -and -not [string]::IsNullOrWhiteSpace($textValue)) {
        return ("<a href='$(ConvertTo-HtmlEncodedText $safeLink)' target='_blank' rel='noopener noreferrer'>$encodedText</a>")
    }
    $pattern = '(https?://[^\s<]+)'
    return [regex]::Replace($encodedText, $pattern, {
        param($match)
        $raw = $match.Value
        $trimmed = $raw.TrimEnd('.', ',', ';', ':', '。', '，', '；', '：')
        $suffix = $raw.Substring($trimmed.Length)
        $decoded = $trimmed -replace '&amp;', '&'
        $detectedLink = Get-SafeHttpUrl -Value $decoded
        if ([string]::IsNullOrWhiteSpace($detectedLink)) { return $raw }
        return "<a href='$(ConvertTo-HtmlEncodedText $detectedLink)' target='_blank' rel='noopener noreferrer'>$trimmed</a>$suffix"
    })
}

function Get-DefaultRuleFields {
    param($Item)
    $itemType = [string]$Item.类型
    $namePattern = [string]$Item.名称
    $version = [string]$Item.版本
    $matchType = if ($itemType -eq '软件') { '精确' } else { '插件ID精确' }

    if ($itemType -eq '软件' -and -not [string]::IsNullOrWhiteSpace($version)) {
        $escapedVersion = [regex]::Escape($version.Trim())
        $suffixPattern = '(?i)\s*(?:[-–—_/]\s*)?(?:v(?:ersion)?\s*)?\(?' + $escapedVersion + '\)?\s*$'
        $baseName = [regex]::Replace($namePattern, $suffixPattern, '').Trim()
        if (-not [string]::IsNullOrWhiteSpace($baseName) -and $baseName -ne $namePattern.Trim()) {
            $namePattern = $baseName
            $matchType = '包含'
        }
    }

    return [PSCustomObject]@{
        MatchType = $matchType
        NamePattern = $namePattern
        Version = $version
    }
}

function Get-RowHtml {
    param($Item, [bool]$NeedsAction)
    $name = ConvertTo-HtmlEncodedText $Item.名称
    $version = ConvertTo-HtmlEncodedText $Item.版本
    $publisherText = ConvertTo-HtmlEncodedText $Item.发布者
    $publisher = if ([string]::IsNullOrWhiteSpace($publisherText)) { '' } else { "<div class='tooltip-wrapper'><div class='truncate-hover'>$publisherText</div><div class='copyable-tooltip'>$publisherText</div></div>" }
        $reason = Get-UserNoteText $Item.原因
    $status = ConvertTo-HtmlEncodedText $Item.状态
    $matchedRule = $Item.MatchedRule
    $defaultRule = Get-DefaultRuleFields -Item $Item
    $defaultNamePattern = ConvertTo-HtmlEncodedText $defaultRule.NamePattern
    $defaultVersion = ConvertTo-HtmlEncodedText $defaultRule.Version
    $noteReason = if ($null -ne $matchedRule) { ConvertTo-HtmlEncodedText $matchedRule.'备注/原因' } else { '' }
    $noteLink = if ($null -ne $matchedRule) { ConvertTo-HtmlEncodedText $matchedRule.'备注/原因链接' } else { '' }
    $addedBy = if ($null -ne $matchedRule) { ConvertTo-HtmlEncodedText $matchedRule.'添加人' } else { '' }
    $addedTime = if ($null -ne $matchedRule) { ConvertTo-HtmlEncodedText $matchedRule.'添加时间' } else { '' }
    $statusDisplay = $status
    if ([string]$Item.状态 -eq '版本变化' -and $null -ne $matchedRule) {
        $effectiveMatchedRule = Get-EffectiveRuleMatchFields -Rule $matchedRule
        $ruleVersion = ConvertTo-HtmlEncodedText $effectiveMatchedRule.Version
        if (-not [string]::IsNullOrWhiteSpace($ruleVersion)) {
            $statusDisplay = "<div class='tooltip-wrapper'><div class='status-compact'>规则 $ruleVersion</div><div class='copyable-tooltip status-tooltip'>状态：版本变化<br>规则版本：$ruleVersion<br>当前版本：$version</div></div>"
        }
    }

    $typeIcon = if ($Item.类型 -eq '软件') { '&#128187;' } elseif ($Item.类型 -eq 'Chromium插件') { '&#127760;' } else { '&#129418;' }
    $iconKey = ConvertTo-HtmlEncodedText $Item.图标键
    $iconHtml = if (-not [string]::IsNullOrWhiteSpace([string]$Item.图标键)) {
        "<span class='item-icon-slot'><img class='item-icon lazy-checksentry-icon' data-icon-key='$iconKey' alt='' loading='lazy'><span class='type-fallback'>$typeIcon</span></span>"
    } else { "<span class='type-fallback'>$typeIcon</span>" }

    $locationText = ''
    if ($Item.类型 -eq '软件') {
        if (-not [string]::IsNullOrWhiteSpace([string]$Item.安装路径)) { $locationText = '安装位置: ' + [string]$Item.安装路径 }
    } elseif ($Item.Locations) {
        $parts = @()
        foreach ($location in @($Item.Locations)) {
            $browser = if ($location.Browser) { [string]$location.Browser } elseif ($Item.类型 -eq 'Firefox插件') { 'Firefox' } else { '' }
            $windowsUser = [string]$location.WindowsUser
            $locationProfile = [string]$location.ProfileName
            if ($windowsUser -and $browser -and $locationProfile) { $parts += ($windowsUser + ' / ' + $browser + ' / ' + $locationProfile) }
            elseif ($browser -and $locationProfile) { $parts += ($browser + ' / ' + $locationProfile) }
            elseif ($browser) { $parts += $browser }
            elseif ($locationProfile) { $parts += $locationProfile }
        }
        $locationText = '安装位置: ' + (($parts | Select-Object -Unique) -join ' · ')
    }
    $locationHtml = if ([string]::IsNullOrWhiteSpace($locationText)) { '' } else { "<div class='tooltip-wrapper'><div class='item-location'>$(ConvertTo-HtmlEncodedText $locationText)</div><div class='copyable-tooltip'>$(ConvertTo-HtmlEncodedText $locationText)</div></div>" }
    $rowClass = switch ([string]$Item.状态) {
        '命中黑名单' { 'row-red row-banned' }
        '待定' { 'row-pending' }
        '版本变化' { 'row-yellow' }
        default { 'row-green' }
    }

    $actionHtml = ''
    if ($NeedsAction) {
        $matchOptions = if ($Item.类型 -eq '软件' -and $defaultRule.MatchType -eq '包含') {
            '<option value="精确">精确</option><option value="包含" selected>包含</option><option value="通配符">通配符</option><option value="正则">正则</option>'
        } elseif ($Item.类型 -eq '软件') {
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
          <label>关键词<input type="text" name="namePattern" value="$defaultNamePattern" maxlength="512"></label>
          <label>版本号（默认当前版本，清空=不锁版本）<input type="text" name="version" value="$defaultVersion" maxlength="2048"><button type="button" class="use-current-version" data-current-version="$version">用当前版本</button></label>
                    <label>发布者<input type="text" name="publisher" value="$publisherText" maxlength="256"></label>
          <label>备注/原因<input type="text" name="note" value="$noteReason" maxlength="2000"><input type="hidden" name="noteLink" value="$noteLink"></label>
          <label>添加人<input type="text" value="$addedBy" readonly></label>
          <label>添加时间<input type="text" value="$addedTime" readonly></label>

          <div class="btnrow"><button type="submit" data-status="允许" class="btn-approve">加入白名单（允许）</button><button type="submit" data-status="禁止" class="btn-ban">加入黑名单（禁止）</button><button type="submit" data-status="待定" class="btn-pending">标记待定</button></div>
        </form></details>
"@
    }
    $checkbox = if ($NeedsAction) { "<input type='checkbox' class='row-chk'>" } else { '' }
    $reasonHtml = ConvertTo-SafeNoteHtml -Text $reason -Link $Item.'备注/原因链接'
    $reasonWrapper = if ([string]::IsNullOrWhiteSpace($reason)) { "<div class='tooltip-wrapper'><div class='truncate-hover'>-</div><div class='copyable-tooltip reason-tooltip'>-</div></div>" } else { "<div class='tooltip-wrapper'><div class='truncate-hover'>$reasonHtml</div><div class='copyable-tooltip reason-tooltip'>$reasonHtml</div></div>" }
    return @"
    <tr class="$rowClass"><td class="chk-cell">$checkbox</td><td class="item-cell"><div class="item-cell-wrapper">$iconHtml<div class="item-main"><div class="item-title"><div class="tooltip-wrapper"><div class="item-name">$name</div><div class="copyable-tooltip">$name</div></div></div>$locationHtml</div></div></td><td>$version</td><td>$publisher</td><td>$statusDisplay</td><td class="reason">$reasonWrapper</td><td>$actionHtml</td></tr>
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
    $tableHead = '<table><thead><tr><th class="chk-cell"></th><th>图标 / 名称</th><th>版本</th><th>发布者</th><th>状态/分类</th><th>备注/原因</th><th>操作</th></tr></thead><tbody>'
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
    $greenRows = ($green | ForEach-Object { Get-RowHtml -Item $_ -NeedsAction $true }) -join "`n"
    $body += "<div class='section-container'><details class='green-section'><summary><h2 class='sec-green'>&#128994; 已匹配（点击展开，共 $($green.Count) 项）</h2><label class='section-select green-select'><input type='checkbox' class='section-chk'> 全选本区</label></summary>$tableHead$greenRows$tableTail</details></div>"
    $cloudState = $script:CloudSyncState
    if ($cloudState.Configured) {
        $cloudColor = if ($cloudState.UsingCache) { '#8a5a00' } elseif ($cloudState.Status -eq '失败') { '#c0392b' } else { '#276749' }
        $cloudLabel = if ($cloudState.UsingCache) { '云端：缓存规则' } elseif ($cloudState.Status -eq '失败') { '云端：同步失败' } else { '云端：已验证' }
        $shortHash = if ([string]::IsNullOrWhiteSpace([string]$cloudState.RuleHash)) { '-' } else { ([string]$cloudState.RuleHash).Substring(0, [Math]::Min(8, ([string]$cloudState.RuleHash).Length)) }
        $cloudTitle = ConvertTo-HtmlEncodedText ([string]$cloudState.Message)
        $cloudStatusHtml = "<span class='tag' style='color:$cloudColor' title='$cloudTitle'>$cloudLabel｜白 $($cloudState.WhiteCount)｜待 $($cloudState.PendingCount)｜黑 $($cloudState.BlackCount)｜$shortHash</span>"
    } else {
        $cloudStatusHtml = "<span class='tag'>云端：未配置</span>"
    }
    $summary = "<div class='summary'><span>核对时间：$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</span><span class='tag tag-red'>黑名单：$($red.Count)</span><span class='tag tag-yellow'>版本变化：$($yellow.Count)</span><span class='tag tag-pending'>待定：$($pending.Count)</span><span class='tag tag-green'>已匹配：$($green.Count)</span>$cloudStatusHtml<span class='tag'>清单文件：$(ConvertTo-HtmlEncodedText $Path)</span><button type='button' id='openManagementButton'>清单维护</button><button type='button' id='cloudSettingsButton'>云端清单设置</button><button type='button' id='cloudSyncButton'>立即同步</button><button type='button' id='reloadButton'>重新扫描分类</button></div>"
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
    $identity = Get-RuleIdentity -ItemType $rule.类型 -ExtensionId $rule.插件ID -NamePattern $rule.软件名关键词 -Publisher $rule.发布者 -Version $rule.版本号
    $originalSheet = [string]$Data.originalSheet
    $originalId = [string]$Data.id
    $hasMetadataInput = ($Data.metadataEdit -eq $true)
    $isExplicitClassification = $ExplicitClassification.IsPresent
    $operation = {
        param($package)
        if (-not [string]::IsNullOrWhiteSpace($originalId)) {
            $null = Assert-AllowedSheet -SheetName $originalSheet
            try {
                $metadata = Get-RuleMetadataByIdFromPackage -Package $package -SheetName $originalSheet -RuleId $originalId
                if (-not $script:TakeoverUnlocked -or -not $hasMetadataInput) {
                    $rule.添加人 = $metadata.添加人
                    $rule.添加时间 = $metadata.添加时间
                }
                Remove-RuleByIdFromPackage -Package $package -SheetName $originalSheet -RuleId $originalId
            } catch {
                # 容错：如果按 ID 找不到，可能是因为多次连续保存前端未刷新最新 ID。
                # 此时尝试通过 Identity 删除旧规则。
                $foundSheet = $null
                $foundId = $null
                foreach ($sheet in $script:AllowedSheets) {
                    $worksheet = $package.Workbook.Worksheets[$sheet]
                    if ($null -eq $worksheet.Dimension) { continue }
                    $headers = Get-ExpectedHeaders -SheetName $sheet
                    $idCol = [array]::IndexOf($headers, 'id') + 1
                    for ($r = 2; $r -le $worksheet.Dimension.End.Row; $r++) {
                        if ((Get-WorksheetRowIdentity -Worksheet $worksheet -Row $r -SheetName $sheet) -eq $identity) {
                            $foundSheet = $sheet
                            $foundId = [string]($worksheet.Cells[$r, $idCol].Value)
                            $metadata = Get-RuleMetadataByIdFromPackage -Package $package -SheetName $sheet -RuleId $foundId
                            if (-not $script:TakeoverUnlocked -or -not $hasMetadataInput) {
                                $rule.添加人 = $metadata.添加人
                                $rule.添加时间 = $metadata.添加时间
                            }
                            Remove-RuleByIdFromPackage -Package $package -SheetName $sheet -RuleId $foundId
                            break
                        }
                    }
                    if ($null -ne $foundId) { break }
                }
            }
        }
        
        # 清理可能存在的同身份其他规则，防止冲突
        Remove-IdentityFromPackage -Package $package -Identity $identity
        return Add-RuleToPackage -Package $package -SheetName $TargetSheet -Rule $rule
    }
    return Invoke-WorkbookTransaction -Path $Path -Operation $operation -CreateUndo
}

function Get-ScanLoadingHtml {
    param([string]$Nonce)
    $html = @'
<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>CheckSentry 正在扫描</title>
<style>body{margin:0;background:#f4f7fb;color:#1f2937;font-family:"Microsoft YaHei",Segoe UI,sans-serif;display:grid;place-items:center;min-height:100vh}.card{width:min(520px,86vw);background:white;border-radius:18px;padding:34px;box-shadow:0 18px 50px rgba(15,23,42,.12)}h1{font-size:24px;margin:0 0 12px}.bar{height:12px;background:#e5e7eb;border-radius:999px;overflow:hidden;margin:24px 0 14px}.fill{height:100%;width:8%;background:linear-gradient(90deg,#2563eb,#22c55e);transition:width .35s}.muted{color:#64748b;font-size:14px}.error{color:#b91c1c;white-space:pre-wrap}</style></head><body><main class="card"><h1>正在核对这台电脑</h1><div class="muted" id="status">软件、Chromium 插件和 Firefox 插件正在后台扫描。</div><div class="bar"><div class="fill" id="fill"></div></div><div class="muted" id="elapsed">请稍候，页面会自动进入报告。</div></main>
<script nonce="__NONCE__">async function poll(){try{const r=await fetch('/api/scan-status',{cache:'no-store'});const d=await r.json();if(d.Error){document.getElementById('status').className='error';document.getElementById('status').textContent=d.Error;return;}const percent=Math.max(8,Math.round((d.Completed/d.Total)*100));document.getElementById('fill').style.width=percent+'%';document.getElementById('status').textContent='已完成 '+d.Completed+' / '+d.Total+' 个扫描源';document.getElementById('elapsed').textContent='已用时 '+d.ElapsedSeconds+' 秒';if(d.Ready){location.replace('/');return;}}catch(e){document.getElementById('status').textContent='正在连接本地扫描服务…';}setTimeout(poll,500)}poll();</script></body></html>
'@
    return $html.Replace('__NONCE__', (ConvertTo-HtmlEncodedText $Nonce))
}

function Show-RuleConflictWarnings {
    param([string]$Path)
    $entries = @()
    $allRules = Import-AllRuleSheets -Path $Path
    foreach ($sheetName in $script:AllowedSheets) {
        foreach ($rule in @($allRules.$sheetName)) {
            $itemType = if ([string]::IsNullOrWhiteSpace([string]$rule.'类型')) { '软件' } else { [string]$rule.'类型' }
            $entries += [PSCustomObject]@{ Identity = Get-RuleIdentity -ItemType $itemType -ExtensionId ([string]$rule.'插件ID') -NamePattern ([string]$rule.'软件名关键词') -Publisher ([string]$rule.'发布者') -Version ([string]$rule.'版本号'); Sheet = $sheetName }
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
    Start-InventoryScanJobs -IncludeSystem:$IncludeSystem -IncludeInactive:$IncludeInactiveExtensions -AllUsers:$ScanAllUsers
    try { Start-Process $url -ErrorAction Stop } catch { Write-Host "无法自动打开浏览器，请手动访问：$url" -ForegroundColor Yellow }

    try {
        while ($listener.IsListening) {
            try { $context = $listener.GetContext() } catch { if (-not $listener.IsListening) { break }; continue }
            $request = $context.Request
            $response = $context.Response
            try {
                $route = $request.Url.AbsolutePath
                if ($request.HttpMethod -eq 'GET' -and $route -eq '/api/scan-status') {
                    Write-JsonResponse -Response $response -Object (Get-InventoryScanStatus) -StatusCode 200 -Nonce $nonce
                } elseif ($request.HttpMethod -eq 'GET' -and $route -eq '/api/icon') {
                    $iconKey = [string]$request.QueryString['key']
                    $iconDataUri = Get-IconDataUriByKey -Key $iconKey
                    Write-JsonResponse -Response $response -Object @{ ok = (-not [string]::IsNullOrWhiteSpace($iconDataUri)); dataUri = $iconDataUri } -StatusCode 200 -Nonce $nonce
                } elseif ($request.HttpMethod -eq 'GET' -and $route -eq '/manage') {
                    $template = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'management_template.html') -Raw -Encoding UTF8
                    $template = $template.Replace('__CSRF_TOKEN__', (ConvertTo-HtmlEncodedText $csrfToken)).Replace('__CSP_NONCE__', (ConvertTo-HtmlEncodedText $nonce))
                    Write-HtmlResponse -Response $response -Html $template -StatusCode 200 -Nonce $nonce
                                } elseif ($request.HttpMethod -eq 'GET' -and $route -eq '/api/rules') {
                    $inventory = Get-InventoryData -IncludeSystem:$IncludeSystem
                    $allRules = Import-AllRuleSheets -Path $Path
                    Prime-RuleIconCacheFromReport -Inventory $inventory -AllRules $allRules -IncludeSystem $IncludeSystem
                    $black = @(Convert-RulesForWeb -Rules @($allRules.'黑名单') -IncludeSystem $IncludeSystem)

                                        $white = @(Convert-RulesForWeb -Rules @($allRules.'白名单') -IncludeSystem $IncludeSystem)
                    $pending = @(Convert-RulesForWeb -Rules @($allRules.'待定') -IncludeSystem $IncludeSystem)
                    Write-JsonResponse -Response $response -Object @{ ok = $true; blackRules = $black; whiteRules = $white; pendingRules = $pending; canUndo = ($null -ne $script:LastUndoSnapshot); takeoverUnlocked = $script:TakeoverUnlocked } -StatusCode 200 -Nonce $nonce

                } elseif ($request.HttpMethod -eq 'GET' -and $route -eq '/api/manage/cloudSettings') {
                    $cloudSettings = Read-CloudSettings
                    $cloudUrl = if ($null -ne $cloudSettings) { [string]$cloudSettings.url } else { '' }
                    Write-JsonResponse -Response $response -Object @{ ok = $true; configured = ($null -ne $cloudSettings); url = $cloudUrl; syncState = $script:CloudSyncState } -StatusCode 200 -Nonce $nonce

                } elseif ($request.HttpMethod -eq 'GET' -and $route -eq '/') {
                    $forceRefresh = $request.QueryString['refresh'] -eq '1'
                    if ($forceRefresh) {
                        try {
                            $cloudSyncResult = Sync-CloudWorkbook -Path $Path
                            if ($cloudSyncResult.Configured -and $cloudSyncResult.Changed) { Write-Host "重新扫描：已从云端同步规则到本地清单：$($cloudSyncResult.Count) 条" -ForegroundColor Green }
                        } catch {
                            throw "重新扫描前云端规则验证失败：$($_.Exception.Message)"
                        }
                        Start-InventoryScanJobs -IncludeSystem:$IncludeSystem -IncludeInactive:$IncludeInactiveExtensions -AllUsers:$ScanAllUsers
                    }
                    $scanStatus = Get-InventoryScanStatus
                    if (-not $scanStatus.Ready) {
                        if (-not [string]::IsNullOrWhiteSpace([string]$scanStatus.Error)) { throw [string]$scanStatus.Error }
                        Write-HtmlResponse -Response $response -Html (Get-ScanLoadingHtml -Nonce $nonce) -StatusCode 200 -Nonce $nonce
                        continue
                    }
                    $inventory = Get-InventoryData -IncludeSystem:$IncludeSystem
                    $allRules = Import-AllRuleSheets -Path $Path
                    $black = @($allRules.'黑名单')
                    $white = @($allRules.'白名单')
                    $pending = @($allRules.'待定')
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
                    if ($route -eq '/api/manage/password') {
                        Set-TakeoverPasswords -Data $data
                        Write-JsonResponse -Response $response -Object @{ ok = $true } -StatusCode 200 -Nonce $nonce
                    } elseif ($route -eq '/api/manage/cloudSettings') {
                        $cloudUrl = Get-LimitedText -Value $data.url -FieldName 'Google Sheets 链接' -MaximumLength 2048 -Required
                        $syncResult = Sync-CloudWorkbook -Path $Path -Url $cloudUrl -RequireRemote -CreateUndo
                        $exportUrl = Write-CloudSettings -Url $cloudUrl
                        Write-JsonResponse -Response $response -Object @{ ok = $true; url = $cloudUrl; exportUrl = $exportUrl; configured = $true; changed = $syncResult.Changed; count = $syncResult.Count; usingCache = $syncResult.UsingCache; message = $syncResult.Message; syncedAt = $syncResult.SyncedAt; whiteCount = $syncResult.WhiteCount; pendingCount = $syncResult.PendingCount; blackCount = $syncResult.BlackCount; ruleHash = $syncResult.RuleHash } -StatusCode 200 -Nonce $nonce
                    } elseif ($route -eq '/api/manage/cloudSync') {
                        $syncResult = Sync-CloudWorkbook -Path $Path -CreateUndo
                        Write-JsonResponse -Response $response -Object @{ ok = $true; configured = $syncResult.Configured; changed = $syncResult.Changed; count = $syncResult.Count; usingCache = $syncResult.UsingCache; message = $syncResult.Message; syncedAt = $syncResult.SyncedAt; whiteCount = $syncResult.WhiteCount; pendingCount = $syncResult.PendingCount; blackCount = $syncResult.BlackCount; ruleHash = $syncResult.RuleHash; canUndo = ($null -ne $script:LastUndoSnapshot) } -StatusCode 200 -Nonce $nonce
                    } elseif ($route -eq '/api/manage/takeover') {
                        $settings = Read-TakeoverSettings
                        if ($null -eq $settings) { throw '尚未设置接管密码，请先点击“设置接管密码”。' }
                        $password = Get-LimitedText -Value $data.password -FieldName '使用密码' -MaximumLength 256 -Required
                        if (-not (Test-PasswordRecord -Password $password -Record $settings.usage)) { throw '使用密码不正确。' }
                        $script:TakeoverUnlocked = $true
                        Write-JsonResponse -Response $response -Object @{ ok = $true; unlocked = $true } -StatusCode 200 -Nonce $nonce
                    } elseif ($route -eq '/api/manage/saveMetadata') {
                        if (-not $script:TakeoverUnlocked) { throw '请先点击“接管记录”并输入使用密码。' }
                        $sheet = Assert-AllowedSheet -SheetName $data.originalSheet
                        $id = Get-LimitedText -Value $data.id -FieldName '规则ID' -MaximumLength 10 -Required
                        $addedBy = Get-LimitedText -Value $data.addedBy -FieldName '添加人' -MaximumLength 128 -Required
                        $addedTime = Get-LimitedText -Value $data.addedTime -FieldName '添加时间' -MaximumLength 128
                        if ([string]::IsNullOrWhiteSpace($addedTime)) { $addedTime = Get-Date -Format 'yyyy-MM-dd HH:mm' }
                        $operation = {
                            param($package)
                            Set-RuleMetadataByIdInPackage -Package $package -SheetName $sheet -RuleId $id -AddedBy $addedBy -AddedTime $addedTime
                        }
                        $null = Invoke-WorkbookTransaction -Path $Path -Operation $operation -CreateUndo
                        Write-JsonResponse -Response $response -Object @{ ok = $true; canUndo = $true; addedBy = $addedBy; addedTime = $addedTime } -StatusCode 200 -Nonce $nonce
                    } elseif ($route -eq '/api/manage/batchMove') {
                        $sourceSheet = Assert-AllowedSheet -SheetName $data.sourceSheet
                        $targetSheet = Assert-AllowedSheet -SheetName $data.targetSheet
                        if ($sourceSheet -eq $targetSheet) { throw '目标清单与当前清单相同，无需移动。' }
                        $rawItems = @($data.items)
                        if ($rawItems.Count -lt 1 -or $rawItems.Count -gt 500) { throw '批量处理必须包含 1 到 500 项。' }
                        $ruleIds = @()
                        $seenRuleIds = @{}
                        foreach ($rawItem in $rawItems) {
                            $ruleId = Get-LimitedText -Value $rawItem.id -FieldName '规则ID' -MaximumLength 10 -Required
                            if ($ruleId -notmatch '^\d{1,10}$') { throw '规则 ID 无效。' }
                            if ($seenRuleIds.ContainsKey($ruleId)) { throw '批量请求包含重复规则。' }
                            $seenRuleIds[$ruleId] = $true
                            $ruleIds += $ruleId
                        }
                        $operation = {
                            param($package)
                            $moveEntries = @()
                            $identitySet = @{}
                            foreach ($ruleId in $ruleIds) {
                                $rule = Get-RuleByIdFromPackage -Package $package -SheetName $sourceSheet -RuleId $ruleId
                                $identity = Get-RuleIdentity -ItemType $rule.类型 -ExtensionId $rule.插件ID -NamePattern $rule.软件名关键词 -Publisher $rule.发布者 -Version $rule.版本号
                                if ($identitySet.ContainsKey($identity)) { throw "批量请求包含重复对象：$identity" }
                                $identitySet[$identity] = $true
                                $moveEntries += [PSCustomObject]@{ Rule = $rule; Identity = $identity }
                            }
                            Remove-IdentitiesFromPackage -Package $package -IdentitySet $identitySet
                            $appendState = Get-WorksheetAppendState -Package $package -SheetName $targetSheet
                            foreach ($entry in $moveEntries) {
                                $null = Add-RuleToPackageWithState -State $appendState -Rule $entry.Rule
                            }
                            return $moveEntries.Count
                        }
                        $count = Invoke-WorkbookTransaction -Path $Path -Operation $operation -CreateUndo
                        Write-JsonResponse -Response $response -Object @{ ok = $true; successCount = $count; canUndo = $true } -StatusCode 200 -Nonce $nonce
                    } elseif ($route -eq '/api/manage/undo') {

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
                            $identity = Get-RuleIdentity -ItemType $rule.类型 -ExtensionId $rule.插件ID -NamePattern $rule.软件名关键词 -Publisher $rule.发布者 -Version $rule.版本号
                            if ($seenIdentities.ContainsKey($identity)) { throw "批量请求中存在重复对象：$identity" }
                            $seenIdentities[$identity] = $true
                            $hasMetadataInput = ($item.metadataEdit -eq $true)
                            $validated += [PSCustomObject]@{ Target = $target; Rule = $rule; Identity = $identity; OriginalSheet = [string]$item.matchedSheet; OriginalId = [string]$item.matchedRuleId; HasMetadataInput = $hasMetadataInput }

                        }
                                                $operation = {
                            param($package)
                            foreach ($entry in $validated) {
                                if (-not [string]::IsNullOrWhiteSpace($entry.OriginalId)) {
                                    $null = Assert-AllowedSheet -SheetName $entry.OriginalSheet
                                    $metadata = Get-RuleMetadataByIdFromPackage -Package $package -SheetName $entry.OriginalSheet -RuleId $entry.OriginalId
                                    if (-not $script:TakeoverUnlocked -or -not $entry.HasMetadataInput) {
                                        $entry.Rule.添加人 = $metadata.添加人
                                        $entry.Rule.添加时间 = $metadata.添加时间
                                    }
                                }
                            }
                            Remove-IdentitiesFromPackage -Package $package -IdentitySet $seenIdentities
                            $appendStates = @{}

                            foreach ($entry in $validated) {
                                if (-not $appendStates.ContainsKey($entry.Target)) {
                                    $appendStates[$entry.Target] = Get-WorksheetAppendState -Package $package -SheetName $entry.Target
                                }
                                $null = Add-RuleToPackageWithState -State $appendStates[$entry.Target] -Rule $entry.Rule
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
        Stop-InventoryScanJobs
        if ($null -ne $listener) {
            try { $listener.Stop() } catch { Write-Verbose "停止 HTTP 监听器失败：$($_.Exception.Message)" }
            try { $listener.Close() } catch { Write-Verbose "关闭 HTTP 监听器失败：$($_.Exception.Message)" }
        }
    }
}

if (-not $LibraryOnly) {
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
    $script:DataDirectory = $parentDirectory
    if (-not (Test-Path -LiteralPath $ListPath -PathType Leaf)) {
        if (-not (Test-Path -LiteralPath $templatePath -PathType Leaf)) { throw '找不到 list_template.xlsx。' }
        Ensure-WorkbookCompatibility -Path $templatePath
        Ensure-CanonicalWorkbookSchema -Path $templatePath
        Assert-WorkbookPathSchema -Path $templatePath
        Copy-Item -LiteralPath $templatePath -Destination $ListPath -ErrorAction Stop
        Write-Host "已从空白模板创建清单：$ListPath" -ForegroundColor Yellow
    }
    Ensure-WorkbookCompatibility -Path $ListPath
    Ensure-CanonicalWorkbookSchema -Path $ListPath
    $clearedSystemNotes = Clear-SystemGeneratedRuleNotes -Path $ListPath
    if ($clearedSystemNotes -gt 0) { Write-Host "已清理系统生成的备注/原因：$clearedSystemNotes 条" -ForegroundColor Yellow }
    Assert-WorkbookPathSchema -Path $ListPath
    try {
        $cloudSyncResult = Sync-CloudWorkbook -Path $ListPath
        if ($cloudSyncResult.Configured -and $cloudSyncResult.UsingCache) { Write-Host $cloudSyncResult.Message -ForegroundColor Yellow }
        elseif ($cloudSyncResult.Configured -and $cloudSyncResult.Changed) { Write-Host "已从云端同步并验证规则：$($cloudSyncResult.Count) 条" -ForegroundColor Green }
        elseif ($cloudSyncResult.Configured) { Write-Host '云端规则已完整验证，本地清单无需更新。' -ForegroundColor DarkGray }
    } catch {
        throw "云端规则已配置，但远程下载和最后成功快照均不可用。为避免产生错误待定结果，本次已停止扫描。详细错误：$($_.Exception.Message)"
    }
    Show-RuleConflictWarnings -Path $ListPath
    Start-ReportServer -RequestedPort $Port -Path $ListPath -IncludeSystem $IncludeSystemComponents.IsPresent
} catch {
    $startupFailureMessage = "CheckSentry 启动失败：$($_.Exception.Message)"
    Write-Host $startupFailureMessage -ForegroundColor Red
    if (-not [string]::IsNullOrWhiteSpace($LogPath)) {
        try {
            $logDirectory = [System.IO.Path]::GetDirectoryName([System.IO.Path]::GetFullPath($LogPath))
            if (-not [string]::IsNullOrWhiteSpace($logDirectory)) { [System.IO.Directory]::CreateDirectory($logDirectory) | Out-Null }
            [System.IO.File]::AppendAllText($LogPath, ((Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + ' ' + $startupFailureMessage + [Environment]::NewLine), $script:Utf8ConsoleEncoding)
        } catch {}
    }
    exit 1
}
}
