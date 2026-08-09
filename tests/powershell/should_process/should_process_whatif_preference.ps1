# vybe-test: powershell/should_process/should_process_whatif_preference
function Test-WhatIfPref {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param()
    if ($PSCmdlet.ShouldProcess("Target")) {
        return "NotSkipped"
    }
    return "WhatIfPrefSkipped"
}
$WhatIfPreference = $true
$res = Test-WhatIfPref
if ($res -ne "WhatIfPrefSkipped") {
    Write-Host "FAIL: \$WhatIfPreference=\$true expected 'WhatIfPrefSkipped', got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
