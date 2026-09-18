# vybe-test: powershell/filesystem_item_cmdlets/get_item_on_directory_returns_directoryinfo
# Get-Item on an existing directory returns a System.IO.DirectoryInfo object
$dir = Join-Path ([System.IO.Path]::GetTempPath()) "vybe_gi_dir_$PID"
New-Item -Path $dir -ItemType Directory -Force | Out-Null

$item = Get-Item $dir

if ($item.GetType().Name -ne "DirectoryInfo") {
    Write-Host "FAIL: expected DirectoryInfo, got $($item.GetType().Name)"
    Remove-Item $dir -Force -Recurse
    exit 1
}

Remove-Item $dir -Force -Recurse
Write-Host "PASS"
exit 0
