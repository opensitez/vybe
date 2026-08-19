# vybe-test: powershell/lexical_escape_rules/escape_octal_sequence
$errors = @()
[void][System.Management.Automation.PSParser]::Tokenize('"`012"', [ref]$errors)
if ($errors.Count -ne 0) {
    Write-Host 'FAIL: octal-style backtick sequence should tokenize as a string token'
    exit 1
}

$value = "`012"
if ($value.Length -ne 4) {
    Write-Host "FAIL: expected a 4-character octal-encoded string, got length $($value.Length)"
    exit 1
}

if ($value[1] -ne [char]0) {
    Write-Host 'FAIL: expected null byte escape at backtick-zero position'
    exit 1
}

Write-Host 'PASS'
exit 0
