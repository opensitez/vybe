# vybe-test: powershell/classes_static_constructors/static_constructor_cannot_have_parameters
# Valid static constructor has no parameter list
class NoParamStatic {
    static [int]$Val
    static NoParamStatic() {
        [NoParamStatic]::Val = 999
    }
}
if ([NoParamStatic]::Val -ne 999) {
    Write-Host "FAIL: NoParamStatic constructor failed"
    exit 1
}
Write-Host "PASS"
exit 0
