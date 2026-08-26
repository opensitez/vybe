# vybe-test: powershell/pipeline_custom_enumerator_streaming/custom_enumerator_is_ienumerable_type_check
class TypeCheckGen : System.Collections.IEnumerable {
    [System.Collections.IEnumerator]GetEnumerator() { return $null }
}
$tc = [TypeCheckGen]::new()
if ($tc -isnot [System.Collections.IEnumerable]) {
    Write-Host "FAIL: Custom class is not recognized as IEnumerable"
    exit 1
}
Write-Host "PASS"
exit 0
