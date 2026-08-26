# vybe-test: powershell/pipeline_chaining/chain_command_success_exit
$log = [System.Collections.Generic.List[string]]::new()
$res = $( $log.Add("First"); $true ) && $( $log.Add("Second"); $true )
if ($log.Count -eq 2 -and $log[0] -eq "First" -and $log[1] -eq "Second") {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
