# vybe-test: powershell/labels/label_in_scriptblock
& { :inner; Write-Output 'PASS' }
exit 0
