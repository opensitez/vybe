# vybe-test: powershell/compare_object_cmdlet/compare_object_property_custom_objects_identical
# -Property compares custom objects by specified key property alone
$obj1 = [pscustomobject]@{ Id = 100; Metadata = "Version 1" }
$obj2 = [pscustomobject]@{ Id = 100; Metadata = "Version 2 (Modified)" }

$diff = Compare-Object @($obj1) @($obj2) -Property Id

if ($null -ne $diff) {
    Write-Host "FAIL: objects with identical Id property expected to compare as equal, got: $diff"
    exit 1
}

Write-Host "PASS"
exit 0
