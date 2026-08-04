# vybe-test: powershell/objects/pscustomobject_add_method
$obj = [PSCustomObject]@{ Name = "Widget"; Price = 9.99 }
$obj | Add-Member -MemberType ScriptMethod -Name "Describe" -Value {
    "$($this.Name) costs `$$($this.Price)"
}
$desc = $obj.Describe()
if ($desc -ne "Widget costs `$9.99") {
    Write-Host "FAIL: '$desc'"
    exit 1
}
Write-Host "PASS"
exit 0
