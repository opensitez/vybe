# vybe-test: powershell/classes_formattable_custom_tostring/formattable_custom_tostring_5
class FormattableItem_5 {
    [int]$Amount = 500
    [string]ToString([string]$format) {
        if ($format -eq "HEX") { return $this.Amount.ToString("X") }
        return $this.Amount.ToString()
    }
}
$item = [FormattableItem_5]::new()
if ($item.ToString("HEX") -ne (500).ToString("X")) { Write-Host "FAIL: Formattable custom ToString failed"; exit 1 }
Write-Host "PASS"; exit 0
