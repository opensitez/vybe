# vybe-test: powershell/select_string_cmdlet/select_string_inputobject_array_parameter
# In PowerShell, passing an array to -InputObject binds the array as a single object (space-concatenated)
$items = @("production-api", "staging-api", "dev-worker")
$matches = @(Select-String -InputObject $items -Pattern "-api")

if ($matches.Count -ne 1) {
    Write-Host "FAIL: expected 1 MatchInfo for non-enumerated -InputObject array, got $($matches.Count)"
    exit 1
}

if ($matches[0].Line -ne "production-api staging-api dev-worker") {
    Write-Host "FAIL: expected space-concatenated line from -InputObject, got: '$($matches[0].Line)'"
    exit 1
}

Write-Host "PASS"
exit 0
