# vybe-test: powershell/pipeline_chaining/chain_in_if_statement
$checked = $false
if (($true) && ($script:checked = $true)) {
    if (-not $checked) {
        Write-Host "FAIL: chain in if condition failed"
        exit 1
    }
}
Write-Host "PASS"
exit 0
