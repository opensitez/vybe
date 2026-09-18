# vybe-test: powershell/filesystem_item_cmdlets/new_item_creates_file_with_itemtype_file
# New-Item -ItemType File creates a new empty file at the specified path
$path = Join-Path ([System.IO.Path]::GetTempPath()) "vybe_ni_file_$PID.txt"

New-Item -Path $path -ItemType File -Force | Out-Null

if (-not (Test-Path $path -PathType Leaf)) {
    Write-Host "FAIL: New-Item file was not created at '$path'"
    exit 1
}

Remove-Item $path -Force
Write-Host "PASS"
exit 0
