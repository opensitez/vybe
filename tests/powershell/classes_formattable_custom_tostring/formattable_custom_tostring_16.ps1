# vybe-test: powershell/classes_formattable_custom_tostring/formattable_custom_tostring_16
class FormattableItem_16 {
    [int]$Amount = 1600
    [string]ToString([string]$format) {
        if ($format -eq "HEX") { return $this.Amount.ToString("X") }
        return $this.Amount.ToString()
    }
}
$item = [FormattableItem_16]::new()
if ($item.ToString("HEX") -ne (1600).ToString("X")) { Write-Host "FAIL: Formattable custom ToString failed"; exit 1 }
Write-Host "PASS"; exit 0
