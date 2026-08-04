# vybe-test: powershell/labels/label_in_for
for ($i = 0; $i -lt 1; $i++) {
    :loop
    Write-Output 'PASS'
}
exit 0
