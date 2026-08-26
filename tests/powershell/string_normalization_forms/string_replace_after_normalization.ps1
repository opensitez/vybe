# vybe-test: powershell/string_normalization_forms/string_replace_after_normalization
$str = "cafe`u{0301} bar".Normalize()
$replaced = $str.Replace("caf`u{00E9}", "tea")
if ($replaced -ne "tea bar") {
    Write-Host "FAIL: String Replace after normalization failed, got $replaced"
    exit 1
}
Write-Host "PASS"
exit 0
