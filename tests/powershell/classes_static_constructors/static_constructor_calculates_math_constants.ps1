# vybe-test: powershell/classes_static_constructors/static_constructor_calculates_math_constants
class MathHelpers {
    static [double]$HalfPi
    static MathHelpers() {
        [MathHelpers]::HalfPi = [math]::PI / 2.0
    }
}
if ([math]::Abs([MathHelpers]::HalfPi - 1.57079632679) -gt 1e-6) {
    Write-Host "FAIL: Static constructor math constant failed"
    exit 1
}
Write-Host "PASS"
exit 0
