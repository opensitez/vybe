# vybe-test: powershell/pipeline_custom_enumerator_streaming/custom_enumerator_measure_object_sum
class ScoreList : System.Collections.IEnumerable {
    [int[]]$Scores = @(10, 20, 30)
    [System.Collections.IEnumerator]GetEnumerator() { return $this.Scores.GetEnumerator() }
}
$sl = [ScoreList]::new()
$meas = $sl | Measure-Object -Sum
if ($meas.Sum -ne 60 -or $meas.Count -ne 3) {
    Write-Host "FAIL: Custom enumerator Measure-Object failed"
    exit 1
}
Write-Host "PASS"
exit 0
