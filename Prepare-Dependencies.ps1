param([string]$PackagePath = '')

$ErrorActionPreference = 'Stop'
$version = '7.8.10'
$destination = Join-Path $PSScriptRoot ('Modules\ImportExcel\' + $version)
$temporaryPackage = $PackagePath
$downloaded = $false
if ([string]::IsNullOrWhiteSpace($temporaryPackage)) {
    $temporaryPackage = Join-Path ([System.IO.Path]::GetTempPath()) ('ImportExcel.' + $version + '.' + [guid]::NewGuid().ToString('N') + '.nupkg')
    Invoke-WebRequest -Uri ('https://www.powershellgallery.com/api/v2/package/ImportExcel/' + $version) -OutFile $temporaryPackage -UseBasicParsing
    $downloaded = $true
}
try {
    if (Test-Path -LiteralPath $destination) { Remove-Item -LiteralPath $destination -Recurse -Force }
    New-Item -ItemType Directory -Path $destination -Force | Out-Null
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($temporaryPackage, $destination)
    $manifest = Join-Path $destination 'ImportExcel.psd1'
    if (-not (Test-Path -LiteralPath $manifest -PathType Leaf)) { throw 'ImportExcel 包结构无效。' }
    $metadata = Test-ModuleManifest -Path $manifest
    if ([version]$metadata.Version -ne [version]$version) { throw "ImportExcel 版本不正确：$($metadata.Version)" }
    $hash = (Get-FileHash -LiteralPath $temporaryPackage -Algorithm SHA256).Hash.ToLowerInvariant()
    [System.IO.File]::WriteAllText((Join-Path $destination 'PACKAGE-SHA256.txt'), ($hash + [Environment]::NewLine), (New-Object System.Text.UTF8Encoding($false)))
    $integrityEntries = @(Get-ChildItem -LiteralPath $destination -Recurse -File | Where-Object { $_.Extension -in @('.ps1','.psm1','.psd1','.dll') } | Sort-Object FullName | ForEach-Object {
        [ordered]@{ path = $_.FullName.Substring($destination.Length + 1).Replace('\','/'); sha256 = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant() }
    })
    $integrityJson = $integrityEntries | ConvertTo-Json -Depth 3
    [System.IO.File]::WriteAllText((Join-Path $destination 'CONTENT-SHA256.json'), $integrityJson, (New-Object System.Text.UTF8Encoding($false)))
    Write-Host "ImportExcel $version 已准备完成。包 SHA-256：$hash" -ForegroundColor Green
} finally {
    if ($downloaded -and (Test-Path -LiteralPath $temporaryPackage)) { Remove-Item -LiteralPath $temporaryPackage -Force }
}
