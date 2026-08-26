# vybe-test: powershell/classes_property_attributes/validateset_on_string_property
class Deployment {
    [ValidateSet("Dev", "Staging", "Prod")][string]$Environment
}
$d = [Deployment]::new()
$d.Environment = "Prod"
$caught = $false
try {
    $d.Environment = "QA"
} catch {
    $caught = $true
}
if ($d.Environment -ne "Prod" -or -not $caught) {
    Write-Host "FAIL: ValidateSet on property failed"
    exit 1
}
Write-Host "PASS"
exit 0
