# vybe-test: powershell/control_flow/switch_no_break_fallthrough
$result = @()
$val = 2
switch ($val) {
    1 { $result += "one" }
    2 { $result += "two" }
    2 { $result += "two-again" }   # switch evaluates all matching arms without break
    3 { $result += "three" }
}
if ($result.Count -ne 2)         { Write-Host "FAIL: count $($result.Count)"; exit 1 }
if ($result[0] -ne "two")        { Write-Host "FAIL: [0]"; exit 1 }
if ($result[1] -ne "two-again")  { Write-Host "FAIL: [1]"; exit 1 }
Write-Host "PASS"
exit 0
