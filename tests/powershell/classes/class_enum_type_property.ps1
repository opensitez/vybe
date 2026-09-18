# vybe-test: powershell/classes/class_enum_type_property
# A PowerShell class can declare properties typed as user-defined enums with default values
enum DeploymentTier {
    Dev
    Staging
    Production
}

class ServiceRecord {
    [string]$ServiceName
    [DeploymentTier]$Tier = [DeploymentTier]::Staging
}

$svc = [ServiceRecord]::new()
$svc.ServiceName = "AuthGateway"

if ($svc.Tier -ne [DeploymentTier]::Staging) {
    Write-Host "FAIL: default enum property value mismatch, got $($svc.Tier)"
    exit 1
}

$svc.Tier = [DeploymentTier]::Production
if ($svc.Tier -ne [DeploymentTier]::Production) {
    Write-Host "FAIL: updated enum property value mismatch, got $($svc.Tier)"
    exit 1
}

Write-Host "PASS"
exit 0
