param(
    [string]$Repo = "zidan-herlangga/universal-media-suite"
)

$ErrorActionPreference = "Stop"

$webDir = Join-Path (Split-Path $PSScriptRoot -Parent) "download"
New-Item -ItemType Directory -Force -Path $webDir | Out-Null

$release = Invoke-RestMethod -Uri "https://api.github.com/repos/$Repo/releases/latest"
Write-Host "Rilis terbaru: $($release.tag_name)"

foreach ($asset in $release.assets) {
    $dest = Join-Path $webDir $asset.name
    Write-Host "Mengunduh: $($asset.name)"
    Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $dest
}

Write-Host "Selesai. File binary siap disajikan oleh situs di folder 'download'."