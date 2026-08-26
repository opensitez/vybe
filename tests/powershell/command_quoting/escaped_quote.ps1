# vybe-test: powershell/command_quoting/escaped_quote
$str = "Quote: `"Hello`""
if ($str -ne 'Quote: "Hello"') {
    Write-Host "FAIL: Escaped quote failed"
    exit 1
}
Write-Host "PASS"
exit 0
