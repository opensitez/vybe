# vybe-test: powershell/argument_transformations/argument_transformation_datetime_parse
class DateParseTransform : System.Management.Automation.ArgumentTransformationAttribute {
    [object] Transform([System.Management.Automation.EngineIntrospector]$e, [object]$i) {
        return [datetime]::ParseExact($i, "yyyyMMdd", $null)
    }
}
function Test-Date {
    param([DateParseTransform()][datetime]$Date)
    return $Date.Year
}
$res = Test-Date "20260807"
if ($res -ne 2026) {
    Write-Host "FAIL: DateParseTransform expected Year=2026, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
