# vybe-test: powershell/parameters_cmdletbinding_supports_should_process/confirm_false_bypasses_confirmation
function Restart-SafeService {
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact="High")]
    param([string]$Svc)
    $ran = $false
    if ($PSCmdlet.ShouldProcess($Svc, "Restart")) {
        $ran = $true
    }
    return $ran
}
$res = Restart-SafeService -Svc "nginx" -Confirm:$false
if ($res -ne $true) {
    Write-Host "FAIL: -Confirm:`$false should bypass confirmation and execute"
    exit 1
}
Write-Host "PASS"
exit 0
