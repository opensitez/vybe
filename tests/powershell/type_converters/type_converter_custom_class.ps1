# vybe-test: powershell/type_converters/type_converter_custom_class
class Point {
    [int]$X
    [int]$Y
}
$pt = [Point]@{ X = 10; Y = 20 }
if ($pt.X -ne 10 -or $pt.Y -ne 20) {
    Write-Host "FAIL: hashtable to custom class [Point] conversion expected X=10, Y=20"
    exit 1
}
Write-Host "PASS"
exit 0
