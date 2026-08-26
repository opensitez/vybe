# vybe-test: powershell/pipeline_measure_object_statistics/measure_character_word_line_statistics
$text = "The quick brown fox`njumps over the lazy dog"
$m = $text | Measure-Object -Character -Word -Line
if ($m.Lines -ne 2 -or $m.Words -ne 9) {
    Write-Host "FAIL: Measure-Object text statistics failed, lines=$($m.Lines), words=$($m.Words)"
    exit 1
}
Write-Host "PASS"
exit 0
