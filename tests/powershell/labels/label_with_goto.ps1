# vybe-test: powershell/labels/label_with_goto
goto label
Write-Host 'FAIL'
exit 1
:label
Write-Host 'PASS'
exit 0
