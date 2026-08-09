# vybe-test: powershell/generic_types/generic_keyvaluepair_type
$kvp = [System.Collections.Generic.KeyValuePair[string, int]]::new("Age", 30)
if ($kvp.Key -ne "Age" -or $kvp.Value -ne 30) {
    Write-Host "FAIL: KeyValuePair[string,int] expected Key='Age', Value=30"
    exit 1
}
Write-Host "PASS"
exit 0
