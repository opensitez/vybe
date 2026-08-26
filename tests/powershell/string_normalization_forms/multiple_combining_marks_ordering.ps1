# vybe-test: powershell/string_normalization_forms/multiple_combining_marks_ordering
$charWithTwoMarks = "c`u{0327}`u{0301}" # c + cedilla + acute
$normC = $charWithTwoMarks.Normalize()
$normD = $normC.Normalize([System.Text.NormalizationForm]::FormD)
if ($normD.Length -ne 3) {
    Write-Host "FAIL: Multiple combining marks decomposition length failed"
    exit 1
}
Write-Host "PASS"
exit 0
