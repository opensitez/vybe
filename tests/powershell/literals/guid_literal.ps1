# vybe-test: powershell/literals/guid_literal
$guid = [guid]::NewGuid()
if ($guid -eq $null) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
