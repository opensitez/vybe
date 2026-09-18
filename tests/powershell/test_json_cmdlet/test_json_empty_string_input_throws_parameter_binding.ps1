# vybe-test: powershell/test_json_cmdlet/test_json_empty_string_input_throws_parameter_binding
# An empty string input fails parameter binding validation because the Json parameter requires non-empty string
$threwBindingError = $false
try {
    '' | Test-Json -ErrorAction Stop
} catch {
    $threwBindingError = ($_.Exception -is [System.Management.Automation.ParameterBindingException])
}

if (-not $threwBindingError) {
    Write-Host "FAIL: piping empty string into Test-Json did not throw ParameterBindingException"
    exit 1
}

Write-Host "PASS"
exit 0
