# vybe-test: powershell/parameters_cmdletbinding_supports_should_process/whatif_preference_variable_global_effect
function Test-GlobalWhatIf {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param([string]$Target)
    if ($PSCmdlet.ShouldProcess($Target)) { return "Ran" }
    return "WhatIf"
}
$oldWhatIf = $WhatIfPreference
try {
    $WhatIfPreference = $true
    $res = Test-GlobalWhatIf -Target "Server1"
    if ($res -ne "WhatIf") {
        Write-Host "FAIL: `$WhatIfPreference=`$true should trigger WhatIf mode"
        exit 1
    }
} finally {
    $WhatIfPreference = $oldWhatIf
}
Write-Host "PASS"
exit 0
