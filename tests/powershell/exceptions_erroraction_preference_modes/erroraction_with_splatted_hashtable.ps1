# vybe-test: powershell/exceptions_erroraction_preference_modes/erroraction_with_splatted_hashtable
function Emit-SplatEA {
    [CmdletBinding()]
    param()
    Write-Error "SplatErr"
}
$p = @{ ErrorAction = "Stop" }
$caught = $false
try {
    Emit-SplatEA @p
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Splatting ErrorAction='Stop' failed"
    exit 1
}
Write-Host "PASS"
exit 0
