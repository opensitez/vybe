# vybe-test: powershell/pipeline_select_object_calculated_properties/mixing_regular_and_calculated_properties
$emp = [pscustomobject]@{ Id = 101; First = "Carol"; Last = "Danvers" }
$res = $emp | Select-Object Id, @{ Name = "Full"; Expression = { "$($_.First) $($_.Last)" } }
if ($res.Id -ne 101 -or $res.Full -ne "Carol Danvers") {
    Write-Host "FAIL: Mixing regular and calculated properties failed"
    exit 1
}
Write-Host "PASS"
exit 0
