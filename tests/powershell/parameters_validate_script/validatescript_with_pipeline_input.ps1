# vybe-test: powershell/parameters_validate_script/validatescript_with_pipeline_input
function Test-ScriptPipe {
    param(
        [Parameter(ValueFromPipeline=$true)]
        [ValidateScript({ $_.Length -ge 4 })]
        [string]$Word
    )
    process { "PIPE:$Word" }
}
$res = "PowerShell" | Test-ScriptPipe
if ($res -ne "PIPE:PowerShell") {
    Write-Host "FAIL: ValidateScript pipeline input failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
