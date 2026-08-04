# vybe-test: powershell/information_streams/information_message_data
Write-Information 'message' -MessageData @{ X = 1 }
Write-Host 'PASS'
exit 0
