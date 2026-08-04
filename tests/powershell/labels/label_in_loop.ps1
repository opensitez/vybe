# vybe-test: powershell/labels/label_in_loop
for ($i=0; $i -lt 1; $i++) {
    :loop
    Write-Output 'PASS'
    break
}
exit 0
