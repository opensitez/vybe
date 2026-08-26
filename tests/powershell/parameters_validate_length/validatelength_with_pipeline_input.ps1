# vybe-test: powershell/parameters_validate_length/validatelength_with_pipeline_input
function Test-LenPipe {
    param(
        [Parameter(ValueFromPipeline=$true)]
        [ValidateLength(4, 10)]
        [string]$Item
    )
    process { "LEN:$($Item.Length)" }
}
$res = "HelloWorld" | Test-LenPipe
if ($res -ne "LEN:10") {
    Write-Host "FAIL: ValidateLength pipeline input failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
