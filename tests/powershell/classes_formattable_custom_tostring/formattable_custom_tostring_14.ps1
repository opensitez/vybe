# vybe-test: powershell/classes_formattable_custom_tostring/formattable_custom_tostring_14
class FormattableItem_14 {
    [int]$Amount = 1400
    [string]ToString([string]$format) {
        if ($format -eq "HEX") { return $this.Amount.ToString("X") }
        return $this.Amount.ToString()
    }
}
$item = [FormattableItem_14]::new()
if ($item.ToString("HEX") -ne (1400).ToString("X")) { Write-Host "FAIL: Formattable custom ToString failed"; exit 1 }
Write-Host "PASS"; exit 0
