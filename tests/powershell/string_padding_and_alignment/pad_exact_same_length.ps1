# vybe-test: powershell/string_padding_and_alignment/pad_exact_same_length
$str = "exact"
$padded = $str.PadLeft(5, [char]'X')
if ($padded -ne "exact") {
    Write-Host "FAIL: PadLeft same length failed"
    exit 1
}
Write-Host "PASS"
exit 0
