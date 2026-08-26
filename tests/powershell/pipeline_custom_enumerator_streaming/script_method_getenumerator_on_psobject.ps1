# vybe-test: powershell/pipeline_custom_enumerator_streaming/script_method_getenumerator_on_psobject
$obj = [pscustomobject]@{ Items = @(10, 20, 30) }
$collected = [System.Collections.Generic.List[int]]::new()
foreach ($item in $obj.Items) {
    $collected.Add($item)
}
if ($collected.Count -ne 3 -or $collected[0] -ne 10) {
    Write-Host "FAIL: ScriptMethod GetEnumerator on PSCustomObject failed"
    exit 1
}
Write-Host "PASS"
exit 0
