# vybe-test: powershell/command_results/multiple_outputs
function Get-Multi {
    return @(10, 20, 30)
}
$res = Get-Multi
if ($res.Length -ne 3 -or $res[1] -ne 20) {
    Write-Host "FAIL: Multiple outputs failed"
    exit 1
}
Write-Host "PASS"
exit 0
