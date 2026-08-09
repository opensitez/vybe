# vybe-test: powershell/type_converters/type_converter_scriptblock_coercion
$sb = [scriptblock]"param(`$x) `$x * 3"
$res = &$sb 7
if ($res -ne 21) {
    Write-Host "FAIL: string to scriptblock conversion expected 21, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
