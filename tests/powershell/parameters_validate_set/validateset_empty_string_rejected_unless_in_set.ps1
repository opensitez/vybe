# vybe-test: powershell/parameters_validate_set/validateset_empty_string_rejected_unless_in_set
function Check-EmptyVal {
    param([ValidateSet("Val1", "Val2")][string]$Val)
    return $Val
}
$caught = $false
try {
    $x = Check-EmptyVal -Val ""
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Empty string should be rejected by ValidateSet"
    exit 1
}
Write-Host "PASS"
exit 0
