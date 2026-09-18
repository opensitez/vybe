# vybe-test: powershell/join_string_cmdlet/join_string_scriptblock_expression
$users = @(
    [pscustomobject]@{ Username = "admin" },
    [pscustomobject]@{ Username = "guest" }
)

# -Property accepts a scriptblock to compute dynamically formatted elements
$res = $users | Join-String -Property { $_.Username.ToUpper() } -Separator '::'

if ($res -ne "ADMIN::GUEST") {
    Write-Host "FAIL: expected 'ADMIN::GUEST', got '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
