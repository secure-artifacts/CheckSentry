param(
    [string]$Version = '1.1.0',
    [ValidateSet('win-x64', 'win-arm64')][string[]]$Runtime = @('win-x64', 'win-arm64'),
    [string]$CertificatePath = '',
    [string]$CertificatePassword = '',
    [string]$OutputDirectory = '',
    [string]$DotNetPath = ''
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) { $OutputDirectory = Join-Path $root 'dist' }
$moduleManifest = Join-Path $root 'Modules\ImportExcel\7.8.10\ImportExcel.psd1'
if (-not (Test-Path -LiteralPath $moduleManifest -PathType Leaf)) { throw '缺少随包 ImportExcel 7.8.10，请先运行 Prepare-Dependencies.ps1。' }
if ([string]::IsNullOrWhiteSpace($DotNetPath)) {
    $dotnetCommand = Get-Command dotnet -ErrorAction SilentlyContinue
    if ($null -eq $dotnetCommand) { throw '未找到 .NET 8 SDK。' }
    $DotNetPath = $dotnetCommand.Source
}
if (-not (Test-Path -LiteralPath $DotNetPath -PathType Leaf)) { throw "dotnet 路径无效：$DotNetPath" }

if (Test-Path -LiteralPath $OutputDirectory) { Remove-Item -LiteralPath $OutputDirectory -Recurse -Force }
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$files = @('README.md','TESTING.md')

foreach ($rid in $Runtime) {
    $publishDirectory = Join-Path $OutputDirectory ('publish-' + $rid)
    $packageDirectory = Join-Path $OutputDirectory ('CheckSentry-' + $Version + '-' + $rid)
    & $DotNetPath publish (Join-Path $root 'launcher\CheckSentry.Launcher.csproj') -c Release -r $rid --self-contained true -p:PublishSingleFile=true -p:Version=$Version -o $publishDirectory
    if ($LASTEXITCODE -ne 0) { throw "Windows 启动器构建失败：$rid" }
    New-Item -ItemType Directory -Path $packageDirectory -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $publishDirectory 'CheckSentry.exe') -Destination $packageDirectory
    foreach ($file in $files) { Copy-Item -LiteralPath (Join-Path $root $file) -Destination $packageDirectory }

    $executable = Join-Path $packageDirectory 'CheckSentry.exe'
    if (-not [string]::IsNullOrWhiteSpace($CertificatePath)) {
        $signTool = Get-Command signtool.exe -ErrorAction SilentlyContinue
        if ($null -eq $signTool) { throw '指定了签名证书，但未找到 Windows SDK signtool.exe。' }
        & $signTool.Source sign /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 /f $CertificatePath /p $CertificatePassword $executable
        if ($LASTEXITCODE -ne 0) { throw "CheckSentry.exe 签名失败：$rid" }
        & $signTool.Source verify /pa /v $executable
        if ($LASTEXITCODE -ne 0) { throw "CheckSentry.exe 签名验证失败：$rid" }
        $securePassword = ConvertTo-SecureString $CertificatePassword -AsPlainText -Force
        $certificate = Get-PfxCertificate -FilePath $CertificatePath -Password $securePassword
    }

    $archivePath = Join-Path $OutputDirectory ('CheckSentry-' + $Version + '-' + $rid + '.zip')
    Compress-Archive -Path (Join-Path $packageDirectory '*') -DestinationPath $archivePath -CompressionLevel Optimal
    $hash = (Get-FileHash -LiteralPath $archivePath -Algorithm SHA256).Hash.ToLowerInvariant()
    [System.IO.File]::WriteAllText(($archivePath + '.sha256'), ($hash + '  ' + [System.IO.Path]::GetFileName($archivePath) + [Environment]::NewLine), (New-Object System.Text.UTF8Encoding($false)))
}

Write-Host "发布包已生成：$OutputDirectory" -ForegroundColor Green
