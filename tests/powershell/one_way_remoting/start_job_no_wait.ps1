# vybe-test: powershell/one_way_remoting/start_job_no_wait
Start-Job -ScriptBlock { Start-Sleep -Milliseconds 1 }
Write-Host "PASS"
exit 0
