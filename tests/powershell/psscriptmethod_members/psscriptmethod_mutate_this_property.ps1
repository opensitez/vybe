# vybe-test: powershell/psscriptmethod_members/psscriptmethod_mutate_this_property
$counter = [pscustomobject]@{ Count = 0 }
$counter | Add-Member -MemberType ScriptMethod -Name "Increment" -Value { $this.Count++ }
$counter.Increment()
$counter.Increment()
if ($counter.Count -ne 2) {
    Write-Host "FAIL: PSScriptMethod mutation expected Count=2, got $($counter.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
