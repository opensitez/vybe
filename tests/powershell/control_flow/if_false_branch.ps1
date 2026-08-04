# vybe-test: powershell/control_flow/if_false_branch
$result = "default"
if ($false) {
    $result = "changed"
}
if ($result -ne "default") {
    Write-Host "FAIL: expected 'default', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
