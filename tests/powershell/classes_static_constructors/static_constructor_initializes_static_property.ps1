# vybe-test: powershell/classes_static_constructors/static_constructor_initializes_static_property
class StaticInitTest {
    static [string]$GlobalGreeting
    static StaticInitTest() {
        [StaticInitTest]::GlobalGreeting = "InitializedStatic"
    }
}
$val = [StaticInitTest]::GlobalGreeting
if ($val -ne "InitializedStatic") {
    Write-Host "FAIL: Static constructor failed, got '$val'"
    exit 1
}
Write-Host "PASS"
exit 0
