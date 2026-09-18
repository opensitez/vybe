# vybe-test: powershell/test_json_cmdlet/test_json_whitespace_only_input_returns_false
# A string containing only whitespace binds to the parameter but fails JSON parsing, returning $false
$res = '   ' | Test-Json -ErrorAction SilentlyContinue

if ($res -ne $false) {
    Write-Host "FAIL: whitespace-only string expected to return `$false, got: $res"
    exit 1
}

Write-Host "PASS"
exit 0
