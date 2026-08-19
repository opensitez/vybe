# vybe-test: powershell/token_and_grammar_system/keyword_true_false
if (($true -is [bool]) -ne $true) {
    Write-Host 'FAIL: $true is not a bool'
    exit 1
}

if ($false -eq $true) {
    Write-Host 'FAIL: $false should not equal $true'
    exit 1
}

if ([string]$true.ToString().ToLower() -ne 'true') {
    Write-Host 'FAIL: boolean true did not stringify to true'
    exit 1
}

Write-Host 'PASS'
exit 0
