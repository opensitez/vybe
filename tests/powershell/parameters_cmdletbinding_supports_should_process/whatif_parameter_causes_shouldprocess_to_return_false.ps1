# vybe-test: powershell/parameters_cmdletbinding_supports_should_process/whatif_parameter_causes_shouldprocess_to_return_false
function Remove-DemoItem {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param([string]$Name)
    $executed = $false
    if ($PSCmdlet.ShouldProcess($Name, "Remove")) {
        $executed = $true
    }
    return $executed
}
$res = Remove-DemoItem -Name "Test" -WhatIf
if ($res -ne $false) {
    Write-Host "FAIL: -WhatIf should make ShouldProcess return false, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
