# vybe-test: powershell/pipeline_measure_object_statistics/measure_ignorewhitespace_text_statistics
$text = "a b c"
$m1 = $text | Measure-Object -Character
$m2 = $text | Measure-Object -Character -IgnoreWhiteSpace
if ($m1.Characters -ne 5 -or $m2.Characters -ne 3) {
    Write-Host "FAIL: Measure-Object -IgnoreWhiteSpace failed, m1=$($m1.Characters), m2=$($m2.Characters)"
    exit 1
}
Write-Host "PASS"
exit 0
