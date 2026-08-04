# vybe-test: powershell/null_coalescing/null_coalescing_command
$value = $null
$result = $value ?? (Get-Command Write-Output)
if (-not $result) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
