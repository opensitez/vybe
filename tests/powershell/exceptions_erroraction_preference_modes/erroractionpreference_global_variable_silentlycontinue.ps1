# vybe-test: powershell/exceptions_erroraction_preference_modes/erroractionpreference_global_variable_silentlycontinue
function Emit-GlobalSilent {
    [CmdletBinding()]
    param()
    Write-Error "SilentGlobal"
    return "OK"
}
$oldEA = $ErrorActionPreference
try {
    $ErrorActionPreference = "SilentlyContinue"
    $res = Emit-GlobalSilent
} finally {
    $ErrorActionPreference = $oldEA
}
if ($res -ne "OK") {
    Write-Host "FAIL: `$ErrorActionPreference='SilentlyContinue' failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
