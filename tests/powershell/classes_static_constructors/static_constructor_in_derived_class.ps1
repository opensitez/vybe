# vybe-test: powershell/classes_static_constructors/static_constructor_in_derived_class
class BaseStaticClass {
    static [string]$BaseMsg
    static BaseStaticClass() {
        [BaseStaticClass]::BaseMsg = "BaseInit"
    }
}
class DerivedStaticClass : BaseStaticClass {
    static [string]$DerivedMsg
    static DerivedStaticClass() {
        [DerivedStaticClass]::DerivedMsg = "DerivedInit"
    }
}
if ([DerivedStaticClass]::DerivedMsg -ne "DerivedInit" -or [BaseStaticClass]::BaseMsg -ne "BaseInit") {
    Write-Host "FAIL: Derived class static constructor failed"
    exit 1
}
Write-Host "PASS"
exit 0
