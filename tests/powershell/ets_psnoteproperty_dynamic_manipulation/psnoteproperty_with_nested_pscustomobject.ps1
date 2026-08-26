# vybe-test: powershell/ets_psnoteproperty_dynamic_manipulation/psnoteproperty_with_nested_pscustomobject
$obj = [pscustomobject]@{
    User = [pscustomobject]@{ Name = "Alice"; Role = "Admin" }
}
if ($obj.User.Name -ne "Alice" -or $obj.User.Role -ne "Admin") {
    Write-Host "FAIL: Nested PSNoteProperty custom object failed"
    exit 1
}
Write-Host "PASS"
exit 0
