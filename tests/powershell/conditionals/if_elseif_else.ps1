# vybe-test: powershell/conditionals/if_elseif_else
if (1 -eq 2) {
    $result = 'a'
} elseif (2 -eq 2) {
    $result = 'b'
} else {
    $result = 'c'
}
if ($result -ne 'b') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
