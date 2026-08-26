# vybe-test: powershell/string_padding_and_alignment/padleft_with_spaces
$str = "42"
$padded = $str.PadLeft(5)
if ($padded -ne "   42" -or $padded.Length -ne 5) {
    Write-Host "FAIL: PadLeft with spaces failed, got '$padded'"
    exit 1
}
Write-Host "PASS"
exit 0
