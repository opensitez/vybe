# vybe-test: powershell/ets_psscriptproperty_getter_and_setter/psscriptproperty_in_pipeline_select_object
$users = @(
    [pscustomobject]@{ F = "A"; L = "1" },
    [pscustomobject]@{ F = "B"; L = "2" }
)
foreach ($u in $users) {
    $u.PSObject.Properties.Add([System.Management.Automation.PSScriptProperty]::new("Combined", { "$($this.F)-$($this.L)" }))
}
$res = @($users | Select-Object -ExpandProperty Combined)
if ($res[0] -ne "A-1" -or $res[1] -ne "B-2") {
    Write-Host "FAIL: PSScriptProperty in Select-Object failed"
    exit 1
}
Write-Host "PASS"
exit 0
