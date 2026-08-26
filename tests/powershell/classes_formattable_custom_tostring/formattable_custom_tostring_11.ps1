# vybe-test: powershell/classes_formattable_custom_tostring/formattable_custom_tostring_11
class FormattableItem_11 {
    [int]$Amount = 1100
    [string]ToString([string]$format) {
        if ($format -eq "HEX") { return $this.Amount.ToString("X") }
        return $this.Amount.ToString()
    }
}
$item = [FormattableItem_11]::new()
if ($item.ToString("HEX") -ne (1100).ToString("X")) { Write-Host "FAIL: Formattable custom ToString failed"; exit 1 }
Write-Host "PASS"; exit 0
