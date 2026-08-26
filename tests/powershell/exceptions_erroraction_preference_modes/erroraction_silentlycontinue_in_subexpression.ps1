# vybe-test: powershell/exceptions_erroraction_preference_modes/erroraction_silentlycontinue_in_subexpression
$res = $(
    function Emit-SubSilent {
        [CmdletBinding()]
        param()
        Write-Error "SubSilent"
        return "SubResult"
    }
    Emit-SubSilent -ErrorAction SilentlyContinue
)
if ($res -ne "SubResult") {
    Write-Host "FAIL: SilentlyContinue in subexpression failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
