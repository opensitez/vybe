# vybe-test: powershell/operators/pipeline_chain_operators
# && runs RHS only if LHS succeeded (exit 0)
# || runs RHS only if LHS failed (exit non-0)
$log = @()
$true  && { $log += "and-true"  } | Out-Null
$false || { $log += "or-false"  } | Out-Null
$false && { $log += "and-false" } | Out-Null   # should NOT run
$true  || { $log += "or-true"   } | Out-Null   # should NOT run
if ($log.Count -ne 2)          { Write-Host "FAIL: count $($log.Count)"; exit 1 }
if ($log[0] -ne "and-true")    { Write-Host "FAIL: [0]"; exit 1 }
if ($log[1] -ne "or-false")    { Write-Host "FAIL: [1]"; exit 1 }
Write-Host "PASS"
exit 0
