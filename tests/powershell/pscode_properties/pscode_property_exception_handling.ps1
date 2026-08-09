# vybe-test: powershell/pscode_properties/pscode_property_exception_handling
class ThrowCodeHelper {
    static [string] GetErr([object]$t) { throw "CodePropertyException" }
}
$g = [ThrowCodeHelper].GetMethod("GetErr")
$obj = [pscustomobject]@{}
$obj | Add-Member -MemberType CodeProperty -Name "ErrProp" -Value $g
try {
    $x = $obj.ErrProp
    Write-Host "FAIL: throwing CodeProperty expected exception"
    exit 1
} catch {
    Write-Host "PASS"
    exit 0
}
