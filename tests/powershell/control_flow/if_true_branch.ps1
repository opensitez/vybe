# vybe-test: powershell/control_flow/if_true_branch
$result = ""
if ($true) {
    $result = "taken"
}
if ($result -ne "taken") {
    Write-Host "FAIL: expected 'taken', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
