# vybe-test: powershell/classes_hidden_members/hidden_flag_inspection_via_reflection
class FlagCheck {
    hidden [string]$HiddenProp = "check"
}
$prop = [FlagCheck].GetProperty("HiddenProp", [System.Reflection.BindingFlags]"NonPublic,Public,Instance")
if ($prop -eq $null) {
    Write-Host "FAIL: Reflection inspection of hidden member failed"
    exit 1
}
Write-Host "PASS"
exit 0
