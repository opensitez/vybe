# vybe-test: powershell/hashtables/hashtable_contains_value
# The ContainsValue method accurately checks if a specific value exists in the hashtable
$registry = @{
    Admin   = "superuser"
    Support = "helpdesk"
    Billing = "finance"
}

$hasFinance = $registry.ContainsValue("finance")
$hasGuest   = $registry.ContainsValue("guest")

if (-not $hasFinance) {
    Write-Host "FAIL: ContainsValue failed to detect existing value 'finance'"
    exit 1
}

if ($hasGuest) {
    Write-Host "FAIL: ContainsValue incorrectly detected missing value 'guest'"
    exit 1
}

Write-Host "PASS"
exit 0
