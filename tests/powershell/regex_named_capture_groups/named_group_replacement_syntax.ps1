# vybe-test: powershell/regex_named_capture_groups/named_group_replacement_syntax
$str = "first: Alice, last: Smith"
$res = $str -replace "first:\s*(?<f>\w+),\s*last:\s*(?<l>\w+)", '${l}, ${f}'
if ($res -ne "Smith, Alice") {
    Write-Host "FAIL: Named group replacement failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
