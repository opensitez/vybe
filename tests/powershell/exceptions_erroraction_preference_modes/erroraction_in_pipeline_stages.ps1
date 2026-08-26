# vybe-test: powershell/exceptions_erroraction_preference_modes/erroraction_in_pipeline_stages
function Step-A { process { 1; 2; 3 } }
function Step-B {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true)][int]$In)
    process {
        if ($In -eq 2) { Write-Error "StepBErr" }
        $In * 10
    }
}
$caught = $false
try {
    Step-A | Step-B -ErrorAction Stop
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: -ErrorAction Stop in pipeline stage failed"
    exit 1
}
Write-Host "PASS"
exit 0
