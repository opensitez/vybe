# vybe-test: powershell/classes_static_constructors/static_constructor_access_triggers_on_static_method_call
class MethodTrigger {
    static [bool]$WasInitialized = $false
    static MethodTrigger() {
        [MethodTrigger]::WasInitialized = $true
    }
    static [string]Echo([string]$s) { return $s }
}
$ret = [MethodTrigger]::Echo("hi")
if (-not [MethodTrigger]::WasInitialized -or $ret -ne "hi") {
    Write-Host "FAIL: Static constructor on static method call failed"
    exit 1
}
Write-Host "PASS"
exit 0
