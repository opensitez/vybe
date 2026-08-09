# vybe-test: powershell/pipeline_chaining/chain_command_failure_exit
$out = @()
(Write-Error "Err" -ErrorAction SilentlyContinue; $global:LASTEXITCODE = 1; $false) || ($script:out += "Fallback")
if ($out.Count -ne 1 -or $out[0] -ne "Fallback") {
    Write-Host "FAIL: command failure || Fallback expected"
    exit 1
}
Write-Host "PASS"
exit 0
