# vybe-test: powershell/pipeline_measure_object_statistics/measure_string_lengths_via_property
$words = @("a", "ab", "abc", "abcd")
$m = $words | Measure-Object -Property Length -Sum -Average -Maximum
if ($m.Sum -ne 10 -or $m.Average -ne 2.5 -or $m.Maximum -ne 4) {
    Write-Host "FAIL: Measure-Object on string Length failed"
    exit 1
}
Write-Host "PASS"
exit 0
