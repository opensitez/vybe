# vybe-test: powershell/ref_parameters/ref_param_scriptblock_invocation
$sb = { param([ref]$r) $r.Value = $r.Value + 10 }
$val = 5
&$sb ([ref]$val)
if ($val -ne 15) {
    Write-Host "FAIL: scriptblock [ref] invocation expected 15, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
