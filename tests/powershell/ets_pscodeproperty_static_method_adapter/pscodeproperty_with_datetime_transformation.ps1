# vybe-test: powershell/ets_pscodeproperty_static_method_adapter/pscodeproperty_with_datetime_transformation
class DateTransformer {
    static [string]GetIso([psobject]$i) { return $i.Date.ToString("yyyy-MM-dd") }
}
$obj = [pscustomobject]@{ Date = [datetime]::Parse("2026-08-26") }
$obj.PSObject.Properties.Add([System.Management.Automation.PSCodeProperty]::new("IsoDate", [DateTransformer].GetMethod("GetIso")))
if ($obj.IsoDate -ne "2026-08-26") {
    Write-Host "FAIL: PSCodeProperty DateTime transformation failed"
    exit 1
}
Write-Host "PASS"
exit 0
