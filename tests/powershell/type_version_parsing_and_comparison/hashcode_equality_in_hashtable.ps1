# vybe-test: powershell/type_version_parsing_and_comparison/hashcode_equality_in_hashtable
$v1 = [version]"1.2.3"
$v2 = [version]"1.2.3"
$ht = @{ $v1 = "ok" }
if ($ht[$v2] -ne "ok") {
    Write-Host "FAIL: Hashtable lookup with identical version failed"
    exit 1
}
Write-Host "PASS"
exit 0
