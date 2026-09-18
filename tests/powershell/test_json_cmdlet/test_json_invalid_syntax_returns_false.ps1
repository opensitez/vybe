# vybe-test: powershell/test_json_cmdlet/test_json_invalid_syntax_returns_false
# Malformed JSON syntax causes Test-Json to return $false
$res = '{"name": "UnterminatedString' | Test-Json -ErrorAction SilentlyContinue

if ($res -ne $false) {
    Write-Host "FAIL: expected `$false for malformed JSON, got: $res"
    exit 1
}

Write-Host "PASS"
exit 0
