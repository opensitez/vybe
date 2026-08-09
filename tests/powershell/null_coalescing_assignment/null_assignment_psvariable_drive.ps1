# vybe-test: powershell/null_coalescing_assignment/null_assignment_psvariable_drive
$variable:nullDriven = $null
$variable:nullDriven ??= "DriveAssigned"
if ($nullDriven -ne "DriveAssigned") {
    Write-Host "FAIL: \$variable: drive ??= expected DriveAssigned, got $nullDriven"
    exit 1
}
Write-Host "PASS"
exit 0
