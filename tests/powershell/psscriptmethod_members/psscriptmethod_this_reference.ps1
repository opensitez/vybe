# vybe-test: powershell/psscriptmethod_members/psscriptmethod_this_reference
$person = [pscustomobject]@{ FirstName = "John"; LastName = "Doe" }
$person | Add-Member -MemberType ScriptMethod -Name "FullName" -Value { "$($this.FirstName) $($this.LastName)" }
$res = $person.FullName()
if ($res -ne "John Doe") {
    Write-Host "FAIL: \$this reference expected 'John Doe', got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
