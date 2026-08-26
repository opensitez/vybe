# vybe-test: powershell/pipeline_nested_steppable_pipeline/steppable_pipeline_exception_handling
$sb = {
    param([Parameter(ValueFromPipeline=$true)][int]$Val)
    process {
        if ($Val -lt 0) { throw "NegativeNotAllowed" }
        $Val
    }
}
$sp = $sb.GetSteppablePipeline()
$sp.Begin($true)
$null = $sp.Process(10)
$caught = $false
try {
    $sp.Process(-5)
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected exception in steppable pipeline Process"
    exit 1
}
Write-Host "PASS"
exit 0
