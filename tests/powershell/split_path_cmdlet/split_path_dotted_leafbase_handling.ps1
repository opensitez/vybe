# vybe-test: powershell/split_path_cmdlet/split_path_dotted_leafbase_handling
# In a file with multiple dots (e.g. app.version.1.0.exe), -LeafBase strips only the final extension
$res = Split-Path "app.version.1.0.exe" -LeafBase

if ($res -ne "app.version.1.0") {
    Write-Host "FAIL: expected 'app.version.1.0', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
