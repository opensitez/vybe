# vybe-test: powershell/ordered_hashtables/ordered_hashtable_to_pscustomobject
$h = [ordered]@{ First = "A"; Second = "B" }
$obj = [pscustomobject]$h
$members = @($obj.psobject.Properties.Name)
if ($members[0] -ne "First" -or $members[1] -ne "Second") {
    Write-Host "FAIL: PSCustomObject property order expected First, Second, got $($members -join ',')"
    exit 1
}
Write-Host "PASS"
exit 0
