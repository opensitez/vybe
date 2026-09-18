# vybe-test: powershell/split_path_cmdlet/split_path_leafbase_without_extension
# Split-Path -LeafBase returns the leaf name stripped of its final file extension
$res = Split-Path "/dir/archive.tar.gz" -LeafBase

if ($res -ne "archive.tar") {
    Write-Host "FAIL: expected 'archive.tar', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
