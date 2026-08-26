# vybe-test: powershell/string_data_parser/case_insensitivity_of_parsed_hashtable
$str = "KeyName = ValueData"
$ht = ConvertFrom-StringData -StringData $str
if ($ht["keyname"] -ne "ValueData" -or $ht["KEYNAME"] -ne "ValueData") {
    Write-Host "FAIL: Case insensitivity of parsed hashtable failed"
    exit 1
}
Write-Host "PASS"
exit 0
