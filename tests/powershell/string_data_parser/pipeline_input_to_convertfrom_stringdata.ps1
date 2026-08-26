# vybe-test: powershell/string_data_parser/pipeline_input_to_convertfrom_stringdata
$str = "k1=v1`nk2=v2"
$ht = $str | ConvertFrom-StringData
if ($ht["k1"] -ne "v1" -or $ht["k2"] -ne "v2") {
    Write-Host "FAIL: Pipeline input to ConvertFrom-StringData failed"
    exit 1
}
Write-Host "PASS"
exit 0
