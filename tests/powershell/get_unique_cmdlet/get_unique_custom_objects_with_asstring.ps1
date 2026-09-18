# vybe-test: powershell/get_unique_cmdlet/get_unique_custom_objects_with_asstring
# With -AsString, custom objects are deduplicated by their string representations
$obj1 = [pscustomobject]@{ Code = "ALPHA" }
$obj2 = [pscustomobject]@{ Code = "ALPHA" }
$obj3 = [pscustomobject]@{ Code = "BETA" }

$unique = @($obj1, $obj2, $obj3 | Get-Unique -AsString)

if ($unique.Count -ne 2) {
    Write-Host "FAIL: expected 2 unique custom objects with -AsString, got $($unique.Count)"
    exit 1
}

Write-Host "PASS"
exit 0
