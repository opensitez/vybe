# vybe-test: powershell/pipeline_select_object_calculated_properties/calculated_property_name_and_expression_hashtable
$users = @(
    [pscustomobject]@{ First = "Alice"; Last = "Smith" },
    [pscustomobject]@{ First = "Bob"; Last = "Jones" }
)
$res = @($users | Select-Object @{ Name = "FullName"; Expression = { "$($_.First) $($_.Last)" } })
if ($res.Length -ne 2 -or $res[0].FullName -ne "Alice Smith" -or $res[1].FullName -ne "Bob Jones") {
    Write-Host "FAIL: Calculated property Name/Expression failed"
    exit 1
}
Write-Host "PASS"
exit 0
