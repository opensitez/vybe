# vybe-test: powershell/pipeline_custom_enumerator_streaming/custom_enumerator_with_group_object
class WordBag : System.Collections.IEnumerable {
    [string[]]$Words = @("cat", "dog", "cow", "duck")
    [System.Collections.IEnumerator]GetEnumerator() { return $this.Words.GetEnumerator() }
}
$wb = [WordBag]::new()
$groups = @($wb | Group-Object -Property Length)
if ($groups.Count -ne 2) { # length 3 (cat, dog, cow) and length 4 (duck)
    Write-Host "FAIL: Custom enumerator Group-Object failed"
    exit 1
}
Write-Host "PASS"
exit 0
