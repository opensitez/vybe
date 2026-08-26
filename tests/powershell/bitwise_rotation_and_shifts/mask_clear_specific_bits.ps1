# vybe-test: powershell/bitwise_rotation_and_shifts/mask_clear_specific_bits
$val = 0xFF
$mask = -bnot 0x0F # clear lower 4 bits
$res = $val -band $mask
if (($res -band 0xFF) -ne 0xF0) {
    Write-Host "FAIL: Mask clear bits failed, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
