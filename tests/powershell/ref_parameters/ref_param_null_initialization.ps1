# vybe-test: powershell/ref_parameters/ref_param_null_initialization
$nullVar = $null
$r = [ref]$nullVar
$r.Value = "Initialized"
if ($nullVar -ne "Initialized") {
    Write-Host "FAIL: null ref variable mutation expected 'Initialized', got $nullVar"
    exit 1
}
Write-Host "PASS"
exit 0
