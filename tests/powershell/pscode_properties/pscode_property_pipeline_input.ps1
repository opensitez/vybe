# vybe-test: powershell/pscode_properties/pscode_property_pipeline_input
class PipeCodeHelper {
    static [int] GetSquared([object]$t) { return $t.Num * $t.Num }
}
$g = [PipeCodeHelper].GetMethod("GetSquared")
$res = 1..3 | ForEach-Object {
    $o = [pscustomobject]@{ Num = $_ }
    $o | Add-Member -MemberType CodeProperty -Name "Sq" -Value $g -PassThru
}
if ($res[0].Sq -ne 1 -or $res[1].Sq -ne 4 -or $res[2].Sq -ne 9) {
    Write-Host "FAIL: pipeline CodeProperty expected Sq 1, 4, 9"
    exit 1
}
Write-Host "PASS"
exit 0
