# vybe-test: powershell/hashtables/hashtable_splatting_parameters_to_function
# Splatting a hashtable with @ prefix passes key-value pairs as named arguments to a function
function FormatEmployeeBadge {
    param(
        [string]$Name,
        [string]$Department,
        [int]$BadgeId
    )
    return "Badge:${BadgeId}-${Name} (${Department})"
}

$badgeData = @{
    Name       = "Dana"
    Department = "Security"
    BadgeId    = 504
}

$badgeString = FormatEmployeeBadge @badgeData

$expected = "Badge:504-Dana (Security)"
if ($badgeString -ne $expected) {
    Write-Host "FAIL: expected '$expected', got '$badgeString'"
    exit 1
}

Write-Host "PASS"
exit 0
