# vybe-test: powershell/filesystem_item_cmdlets/new_item_creates_directory_with_itemtype_directory
# New-Item -ItemType Directory creates a new directory at the specified path
$dir = Join-Path ([System.IO.Path]::GetTempPath()) "vybe_ni_dir_$PID"

New-Item -Path $dir -ItemType Directory -Force | Out-Null

if (-not (Test-Path $dir -PathType Container)) {
    Write-Host "FAIL: New-Item directory was not created at '$dir'"
    exit 1
}

Remove-Item $dir -Force -Recurse
Write-Host "PASS"
exit 0
