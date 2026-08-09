# vybe-test: powershell/psalias_properties/psalias_property_pipeline_input
$res = 1..2 | ForEach-Object {
    $o = [pscustomobject]@{ Val = $_ }
    $o | Add-Member -MemberType AliasProperty -Name "V" -Value "Val" -PassThru
}
if ($res[0].V -ne 1 -or $res[1].V -ne 2) {
    Write-Host "FAIL: pipeline AliasProperty expected V=1, 2"
    exit 1
}
Write-Host "PASS"
exit 0
