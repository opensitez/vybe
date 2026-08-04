# vybe-test: powershell/documents/here_string_documentation
$doc = @'
This is a multi-line
PowerShell here-string document.
'@
if ($doc -notmatch 'PowerShell here-string') {
    Write-Host "FAIL: expected here-string text"
    exit 1
}
Write-Host "PASS"
exit 0
