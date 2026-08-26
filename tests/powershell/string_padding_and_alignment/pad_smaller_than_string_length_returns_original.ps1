# vybe-test: powershell/string_padding_and_alignment/pad_smaller_than_string_length_returns_original
$str = "hello"
$p1 = $str.PadLeft(3)
$p2 = $str.PadRight(3)
if ($p1 -ne "hello" -or $p2 -ne "hello") {
    Write-Host "FAIL: Pad smaller than length should return original string"
    exit 1
}
Write-Host "PASS"
exit 0
