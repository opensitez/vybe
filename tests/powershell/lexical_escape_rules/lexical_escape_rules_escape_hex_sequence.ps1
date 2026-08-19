# vybe-test: powershell/lexical_escape_rules/escape_hex_sequence
$errors = @()
[void][System.Management.Automation.PSParser]::Tokenize('"\x41"', [ref]$errors)
if ($errors.Count -ne 0) {
    Write-Host 'PASS'
    exit 0
}

$val = "\x41"
if ($val -ne '\\x41') {
    Write-Host "FAIL: expected literal hex sequence token text, got $val"
    exit 1
}

Write-Host 'PASS'
exit 0
