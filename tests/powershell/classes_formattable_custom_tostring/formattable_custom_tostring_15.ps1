# vybe-test: powershell/classes_formattable_custom_tostring/formattable_custom_tostring_15
class FormattableItem_15 {
    [int]$Amount = 1500
    [string]ToString([string]$format) {
        if ($format -eq "HEX") { return $this.Amount.ToString("X") }
        return $this.Amount.ToString()
    }
}
$item = [FormattableItem_15]::new()
if ($item.ToString("HEX") -ne (1500).ToString("X")) { Write-Host "FAIL: Formattable custom ToString failed"; exit 1 }
Write-Host "PASS"; exit 0
