# בניית MosquiOrder: בדיקות -> סריקת אבטחה -> בנייה -> ZIP + טביעות SHA-256
# הרצה:  powershell -ExecutionPolicy Bypass -File build.ps1
$ErrorActionPreference = "Stop"
$env:PYTHONUTF8 = "1"
Set-Location $PSScriptRoot

function Step($name, [scriptblock]$cmd) {
    Write-Host "`n=== $name" -ForegroundColor Cyan
    & $cmd
    if ($LASTEXITCODE -ne 0) { throw "$name failed ($LASTEXITCODE)" }
}

Step "Unit tests"        { python -m pytest -q tests }
Step "Bandit (SAST)"     { python -m bandit -q -r core.py app.py }
Step "pip-audit (CVEs)"  { python -m pip_audit -r requirements.txt }
Step "PyInstaller"       { python -m PyInstaller --clean --noconfirm mosquiorder.spec }
Step "Self-test of exe"  {
    $p = Start-Process -FilePath "dist\MosquiOrder\MosquiOrder.exe" -ArgumentList "--selftest" -Wait -PassThru
    $global:LASTEXITCODE = $p.ExitCode
}

$version = (Select-String -Path core.py -Pattern 'APP_VERSION = "(.+)"').Matches[0].Groups[1].Value
Copy-Item columns.example.json dist\MosquiOrder\ -Force
$zip = "dist\MosquiOrder-$version.zip"
if (Test-Path $zip) { Remove-Item $zip }
Compress-Archive -Path dist\MosquiOrder -DestinationPath $zip

$exe = Get-FileHash dist\MosquiOrder\MosquiOrder.exe -Algorithm SHA256
$z = Get-FileHash $zip -Algorithm SHA256
$lines = @(
    "$($exe.Hash.ToLower())  MosquiOrder.exe",
    "$($z.Hash.ToLower())  MosquiOrder-$version.zip"
)
$all = Get-ChildItem dist\MosquiOrder -Recurse -File | Sort-Object FullName | ForEach-Object {
    $h = Get-FileHash $_.FullName -Algorithm SHA256
    "$($h.Hash.ToLower())  $($_.FullName.Substring((Resolve-Path dist\MosquiOrder).Path.Length + 1))"
}
Set-Content -Path dist\SHA256SUMS.txt -Value ($lines + "" + "# all files in MosquiOrder folder:" + $all) -Encoding utf8
Write-Host "`nDone. $zip" -ForegroundColor Green
$lines | ForEach-Object { Write-Host $_ }
