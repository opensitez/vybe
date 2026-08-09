# vybe-test: powershell/using_statements/using_namespace_static_method
using namespace System.IO

$combined = [Path]::Combine("folder", "file.txt")
if ($combined -ne "folder/file.txt" -and $combined -ne "folder\file.txt") {
    Write-Host "FAIL: [Path]::Combine via using namespace expected folder/file.txt, got $combined"
    exit 1
}
Write-Host "PASS"
exit 0
