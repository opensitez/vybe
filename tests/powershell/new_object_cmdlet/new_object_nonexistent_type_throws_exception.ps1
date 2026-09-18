# vybe-test: powershell/new_object_cmdlet/new_object_nonexistent_type_throws_exception
# New-Object throws an exception when given a type name that does not exist
$threw = $false
try {
    New-Object "Completely.Fake.Type.That.Does.Not.Exist" -ErrorAction Stop
} catch {
    $threw = $true
}

if (-not $threw) {
    Write-Host "FAIL: non-existent type did not throw an exception"
    exit 1
}

Write-Host "PASS"
exit 0
