# vybe-test: powershell/pipeline_custom_enumerator_streaming/custom_enumerator_with_select_object_unique
class DupGen : System.Collections.IEnumerable {
    [string[]]$Tags = @("a", "b", "a", "c", "b")
    [System.Collections.IEnumerator]GetEnumerator() { return $this.Tags.GetEnumerator() }
}
$dg = [DupGen]::new()
$uniques = @($dg | Select-Object -Unique)
if ($uniques.Length -ne 3 -or $uniques[0] -ne "a" -or $uniques[2] -ne "c") {
    Write-Host "FAIL: Custom enumerator Select-Object -Unique failed"
    exit 1
}
Write-Host "PASS"
exit 0
