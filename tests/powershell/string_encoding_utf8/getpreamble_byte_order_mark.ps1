# vybe-test: powershell/string_encoding_utf8/getpreamble_byte_order_mark
$encWithBom = [System.Text.UTF8Encoding]::new($true)
$bom = $encWithBom.GetPreamble()
if ($bom.Length -ne 3 -or $bom[0] -ne 0xEF -or $bom[1] -ne 0xBB -or $bom[2] -ne 0xBF) {
    Write-Host "FAIL: UTF8 BOM preamble mismatch"
    exit 1
}
Write-Host "PASS"
exit 0
