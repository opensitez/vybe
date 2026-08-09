# vybe-test: powershell/type_accelerators/type_accelerator_scriptblock
$sb = [scriptblock]::Create("param(`$x) `$x * 2")
$res = &$sb 21
if ($res -ne 42) {
    Write-Host "FAIL: scriptblock execution expected 42, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
