# vybe-test: powershell/ordered_hashtables/ordered_hashtable_nested
$parent = [ordered]@{
    Child = [ordered]@{ Step = "Init"; Status = "OK" }
}
if ($parent.Child.Step -ne "Init") {
    Write-Host "FAIL: nested ordered hashtable expected Init, got $($parent.Child.Step)"
    exit 1
}
Write-Host "PASS"
exit 0
