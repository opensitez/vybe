# vybe-test: powershell/string_normalization_forms/combining_diacritical_mark_character_count
$decomposed = "n`u{0303}" # n + tilde
if ($decomposed.Length -ne 2) {
    Write-Host "FAIL: Decomposed string length expected 2, got $($decomposed.Length)"
    exit 1
}
$composed = $decomposed.Normalize()
if ($composed.Length -ne 1 -or $composed -ne "`u{00F1}") {
    Write-Host "FAIL: Composed n-tilde length expected 1, got $($composed.Length)"
    exit 1
}
Write-Host "PASS"
exit 0
