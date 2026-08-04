# vybe-test: powershell/control_flow/if_else_false
$result = ""
if ($false) {
    $result = "if-branch"
} else {
    $result = "else-branch"
}
if ($result -ne "else-branch") {
    Write-Host "FAIL: expected 'else-branch', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
