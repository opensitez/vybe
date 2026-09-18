# vybe-test: powershell/hashtables/hashtable_case_insensitivity_default
# Default PowerShell hashtables perform case-insensitive key lookups
$data = @{
    ServiceIdentifier = "AuthSvc"
    PortNumber = 8080
}

# Accessing with differing letter case should retrieve the same values
if ($data["serviceidentifier"] -ne "AuthSvc") {
    Write-Host "FAIL: lowercase bracket lookup failed"
    exit 1
}

if ($data.portnumber -ne 8080) {
    Write-Host "FAIL: lowercase dot-property lookup failed"
    exit 1
}

if ($data["SERVICEIDENTIFIER"] -ne "AuthSvc") {
    Write-Host "FAIL: uppercase bracket lookup failed"
    exit 1
}

Write-Host "PASS"
exit 0
