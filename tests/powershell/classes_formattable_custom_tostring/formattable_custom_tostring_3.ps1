# vybe-test: powershell/classes_formattable_custom_tostring/formattable_custom_tostring_3
class FormattableItem_3 {
    [int]$Amount = 300
    [string]ToString([string]$format) {
        if ($format -eq "HEX") { return $this.Amount.ToString("X") }
        return $this.Amount.ToString()
    }
}
$item = [FormattableItem_3]::new()
if ($item.ToString("HEX") -ne (300).ToString("X")) { Write-Host "FAIL: Formattable custom ToString failed"; exit 1 }
Write-Host "PASS"; exit 0
