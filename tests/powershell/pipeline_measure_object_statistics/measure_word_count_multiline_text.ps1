# vybe-test: powershell/pipeline_measure_object_statistics/measure_word_count_multiline_text
$multiline = @("Line one with four", "Line two with four")
$m = $multiline | Measure-Object -Word -Line
if ($m.Words -ne 8 -or $m.Lines -ne 2) {
    Write-Host "FAIL: Multiline array word/line measure failed"
    exit 1
}
Write-Host "PASS"
exit 0
