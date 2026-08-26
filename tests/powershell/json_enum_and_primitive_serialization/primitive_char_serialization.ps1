# vybe-test: powershell/json_enum_and_primitive_serialization/primitive_char_serialization
[char]$c = 'X'
$json = @{ CharVal = $c } | ConvertTo-Json
$recovered = $json | ConvertFrom-Json
if ($recovered.CharVal -ne "X" -and $recovered.CharVal -ne 88) {
    Write-Host "FAIL: Char serialization failed, got '$($recovered.CharVal)'"
    exit 1
}
Write-Host "PASS"
exit 0
