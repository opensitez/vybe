# vybe-test: powershell/objects/pscustomobject_noteproperties
$obj = [PSCustomObject]@{}
$obj | Add-Member -MemberType NoteProperty -Name "Color" -Value "red"
$obj | Add-Member -MemberType NoteProperty -Name "Size"  -Value 42
if ($obj.Color -ne "red") { Write-Host "FAIL: Color"; exit 1 }
if ($obj.Size  -ne 42)    { Write-Host "FAIL: Size";  exit 1 }
$props = $obj | Get-Member -MemberType NoteProperty
if ($props.Count -ne 2) { Write-Host "FAIL: prop count $($props.Count)"; exit 1 }
Write-Host "PASS"
exit 0
