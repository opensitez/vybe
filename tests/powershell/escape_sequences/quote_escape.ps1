# vybe-test: powershell/escape_sequences/quote_escape
# PowerShell escapes a double quote with a BACKTICK. `\"` is not an escape —
# backslash is an ordinary character — and real pwsh fails to parse it.
$s = "He said `"Hi`""
if ($s -eq 'He said "Hi"') { Write-Host 'PASS'; exit 0 }
Write-Host "FAIL: got [$s]"
exit 1
