# vybe-test: powershell/interpolation/boolean_capitalized
# PowerShell renders booleans capitalized — `True`/`False`, not the ECMA
# lowercase spelling.
$text = "$true/$false"
if ($text -ne 'True/False') {
    Write-Host "FAIL: got [$text]"
    exit 1
}
Write-Host 'PASS'
exit 0
