$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$PackageName = "RideReport_Offline_App"
$Dist = Join-Path $Root "dist"
$PackageDir = Join-Path $Dist $PackageName
$ZipPath = Join-Path $Dist "$PackageName.zip"

if (Test-Path $PackageDir) {
  Remove-Item -LiteralPath $PackageDir -Recurse -Force
}
New-Item -ItemType Directory -Path $PackageDir | Out-Null
New-Item -ItemType Directory -Path (Join-Path $PackageDir "app_data\outputs") -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $PackageDir "app_data\uploads") -Force | Out-Null

$Files = @(
  "ride_report_app.py",
  "ride_report_tool.py",
  "Start_RideReport_Offline.bat",
  "offline_requirements.txt",
  "OFFLINE_APP_README.md",
  "LICENSE"
)

foreach ($File in $Files) {
  Copy-Item -LiteralPath (Join-Path $Root $File) -Destination (Join-Path $PackageDir $File) -Force
}

if (Test-Path $ZipPath) {
  Remove-Item -LiteralPath $ZipPath -Force
}
Compress-Archive -LiteralPath $PackageDir -DestinationPath $ZipPath -Force

Write-Host "Created offline package:"
Write-Host $ZipPath
