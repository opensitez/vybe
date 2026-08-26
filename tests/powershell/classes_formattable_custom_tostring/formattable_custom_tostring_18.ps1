# vybe-test: powershell/classes_formattable_custom_tostring/formattable_custom_tostring_18
class FormattableItem_18 {
    [int]$Amount = 1800
    [string]ToString([string]$format) {
        if ($format -eq "HEX") { return $this.Amount.ToString("X") }
        return $this.Amount.ToString()
    }
}
$item = [FormattableItem_18]::new()
if ($item.ToString("HEX") -ne (1800).ToString("X")) { Write-Host "FAIL: Formattable custom ToString failed"; exit 1 }
Write-Host "PASS"; exit 0
