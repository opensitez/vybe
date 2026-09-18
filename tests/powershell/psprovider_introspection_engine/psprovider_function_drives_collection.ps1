# vybe-test: powershell/psprovider_introspection_engine/psprovider_function_drives_collection
# Function provider exposes the 'Function' PSDrive in its Drives collection
$funcP = Get-PSProvider Function

if ($funcP.Name -ne "Function") {
    Write-Host "FAIL: expected provider Name 'Function', got: '$($funcP.Name)'"
    exit 1
}

$driveNames = @($funcP.Drives | ForEach-Object { $_.Name })
if ($driveNames -notcontains "Function") {
    Write-Host "FAIL: 'Function' drive missing from Function provider Drives"
    exit 1
}

Write-Host "PASS"
exit 0
