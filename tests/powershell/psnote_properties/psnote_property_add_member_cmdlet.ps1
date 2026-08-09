# vybe-test: powershell/psnote_properties/psnote_property_add_member_cmdlet
$obj = [pscustomobject]@{}
Add-Member -InputObject $obj -MemberType NoteProperty -Name "DynamicKey" -Value 777
if ($obj.DynamicKey -ne 777) {
    Write-Host "FAIL: Add-Member -InputObject NoteProperty expected DynamicKey=777"
    exit 1
}
Write-Host "PASS"
exit 0
