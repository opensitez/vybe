# vybe-test: powershell/string_padding_and_alignment/padright_with_spaces
$str = "data"
$padded = $str.PadRight(8)
if ($padded -ne "data    " -or $padded.Length -ne 8) {
    Write-Host "FAIL: PadRight with spaces failed, got '$padded'"
    exit 1
}
Write-Host "PASS"
exit 0
