# vybe-test: powershell/compare_object_cmdlet/compare_object_custom_object_passthru_preserves_properties
# -Property with -PassThru outputs original custom objects with their distinct NoteProperties intact
$person1 = [pscustomobject]@{ Name = "Alice"; Role = "Admin" }
$person2 = [pscustomobject]@{ Name = "Bob"; Role = "User" }

$diff = Compare-Object @($person1) @($person2) -Property Name -PassThru

if ($diff.Count -ne 2) {
    Write-Host "FAIL: expected 2 custom object records, got $($diff.Count)"
    exit 1
}

$roles = @($diff | ForEach-Object { $_.Role })
if ($roles -notcontains "Admin" -or $roles -notcontains "User") {
    Write-Host "FAIL: custom object properties not preserved in PassThru: @($($roles -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
