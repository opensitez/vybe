# vybe-test: powershell/ets_psnoteproperty_dynamic_manipulation/psnoteproperty_in_string_interpolation
$obj = [pscustomobject]@{ User = "Bob" }
$msg = "Hello $($obj.User)!"
if ($msg -ne "Hello Bob!") {
    Write-Host "FAIL: PSNoteProperty string interpolation failed"
    exit 1
}
Write-Host "PASS"
exit 0
