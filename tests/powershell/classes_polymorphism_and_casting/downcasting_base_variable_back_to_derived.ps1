# vybe-test: powershell/classes_polymorphism_and_casting/downcasting_base_variable_back_to_derived
class BaseMsg { [string]$Text = "base" }
class SpecialMsg : BaseMsg { [string]$Extra = "extra" }
$s = [SpecialMsg]::new()
[BaseMsg]$b = $s
$downcast = [SpecialMsg]$b
if ($downcast.Extra -ne "extra") {
    Write-Host "FAIL: Downcasting failed"
    exit 1
}
Write-Host "PASS"
exit 0
