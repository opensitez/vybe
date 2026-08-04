# vybe-test: powershell/statement_terminators/newline_in_pipeline
Write-Output 'PASS' |
ForEach-Object { Write-Output $_ }
exit 0
