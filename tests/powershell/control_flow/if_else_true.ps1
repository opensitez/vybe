# vybe-test: powershell/control_flow/if_else_true
$result = ""
if ($true) {
    $result = "if-branch"
} else {
    $result = "else-branch"
}
if ($result -ne "if-branch") {
    Write-Host "FAIL: expected 'if-branch', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
