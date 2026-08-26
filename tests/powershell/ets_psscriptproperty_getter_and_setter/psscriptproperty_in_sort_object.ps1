# vybe-test: powershell/ets_psscriptproperty_getter_and_setter/psscriptproperty_in_sort_object
$items = @(
    [pscustomobject]@{ Str = "ccc" },
    [pscustomobject]@{ Str = "a" },
    [pscustomobject]@{ Str = "bb" }
)
foreach ($it in $items) {
    $it.PSObject.Properties.Add([System.Management.Automation.PSScriptProperty]::new("Len", { $this.Str.Length }))
}
$sorted = @($items | Sort-Object -Property Len)
if ($sorted[0].Str -ne "a" -or $sorted[2].Str -ne "ccc") {
    Write-Host "FAIL: PSScriptProperty in Sort-Object failed"
    exit 1
}
Write-Host "PASS"
exit 0
