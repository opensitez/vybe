# vybe-test: powershell/should_process/should_process_whatif_parameter
function Start-Task {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param([string]$TaskName)
    if ($PSCmdlet.ShouldProcess($TaskName, "Start")) {
        return "Started"
    }
    return "WhatIfSkipped"
}
$res = Start-Task "CleanLog" -WhatIf
if ($res -ne "WhatIfSkipped") {
    Write-Host "FAIL: ShouldProcess with -WhatIf expected 'WhatIfSkipped', got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
