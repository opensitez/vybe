# vybe-test: powershell/scope/using_scope_in_scriptblock
$base = 100
$result = Invoke-Command -ScriptBlock { $using:base + 42 }
if ($result -ne 142) {
    Write-Host "FAIL: expected 142, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
