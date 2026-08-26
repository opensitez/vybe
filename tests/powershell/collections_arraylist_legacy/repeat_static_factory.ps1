# vybe-test: powershell/collections_arraylist_legacy/repeat_static_factory
$al = [System.Collections.ArrayList]::Repeat("echo", 4)
if ($al.Count -ne 4 -or $al[0] -ne "echo" -or $al[3] -ne "echo") {
    Write-Host "FAIL: Repeat factory failed"
    exit 1
}
Write-Host "PASS"
exit 0
