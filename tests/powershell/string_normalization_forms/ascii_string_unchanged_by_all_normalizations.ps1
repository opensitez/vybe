# vybe-test: powershell/string_normalization_forms/ascii_string_unchanged_by_all_normalizations
$ascii = "Powershell 7.4"
$c = $ascii.Normalize([System.Text.NormalizationForm]::FormC)
$d = $ascii.Normalize([System.Text.NormalizationForm]::FormD)
$kc = $ascii.Normalize([System.Text.NormalizationForm]::FormKC)
$kd = $ascii.Normalize([System.Text.NormalizationForm]::FormKD)
if ($c -ne $ascii -or $d -ne $ascii -or $kc -ne $ascii -or $kd -ne $ascii) {
    Write-Host "FAIL: ASCII string mutated by normalization"
    exit 1
}
Write-Host "PASS"
exit 0
