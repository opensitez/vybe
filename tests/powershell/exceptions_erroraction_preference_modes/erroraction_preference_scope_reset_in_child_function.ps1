# vybe-test: powershell/exceptions_erroraction_preference_modes/erroraction_preference_scope_reset_in_child_function
function Outer-ScopeFunc {
    $ErrorActionPreference = "Stop"
    Inner-ScopeFunc
    return $ErrorActionPreference
}
function Inner-ScopeFunc {
    $ErrorActionPreference = "SilentlyContinue"
}
$res = Outer-ScopeFunc
if ($res -ne "Stop") {
    Write-Host "FAIL: Child function should not mutate caller's `$ErrorActionPreference"
    exit 1
}
Write-Host "PASS"
exit 0
