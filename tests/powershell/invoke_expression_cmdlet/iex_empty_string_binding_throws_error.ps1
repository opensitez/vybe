# vybe-test: powershell/invoke_expression_cmdlet/iex_empty_string_binding_throws_error
# Invoke-Expression requires a non-empty Command parameter; an empty string produces an error
$threwError = $false
try {
    Invoke-Expression "" -ErrorAction Stop
} catch {
    $threwError = $true
}

if (-not $threwError) {
    Write-Host "FAIL: empty string did not produce an error"
    exit 1
}

Write-Host "PASS"
exit 0
