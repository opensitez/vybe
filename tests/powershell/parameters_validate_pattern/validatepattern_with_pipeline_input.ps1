# vybe-test: powershell/parameters_validate_pattern/validatepattern_with_pipeline_input
function Test-PatternPipe {
    param(
        [Parameter(ValueFromPipeline=$true)]
        [ValidatePattern('^v\d+\.\d+$')]
        [string]$Ver
    )
    process { "VER:$Ver" }
}
$res = "v1.0" | Test-PatternPipe
if ($res -ne "VER:v1.0") {
    Write-Host "FAIL: ValidatePattern pipeline input failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
