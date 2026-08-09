# vybe-test: powershell/argument_completers/argument_completer_case_insensitivity
$filter = {
    param($w)
    @("Alpha", "BETA") | Where-Object { $_ -like "$w*" }
}
$res = @(&$filter "alpha")
if ($res.Count -ne 1 -or $res[0] -ne "Alpha") {
    Write-Host "FAIL: case-insensitive completer prefix match expected Alpha"
    exit 1
}
Write-Host "PASS"
exit 0
