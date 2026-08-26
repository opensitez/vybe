# vybe-test: powershell/string_padding_and_alignment/trimstart_character_array
$str = "---***hello"
$chars = [char[]]@('-', '*')
$res = $str.TrimStart($chars)
if ($res -ne "hello") {
    Write-Host "FAIL: TrimStart char array failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
