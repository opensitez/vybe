# vybe-test: powershell/classes_formattable_custom_tostring/formattable_custom_tostring_7
class FormattableItem_7 {
    [int]$Amount = 700
    [string]ToString([string]$format) {
        if ($format -eq "HEX") { return $this.Amount.ToString("X") }
        return $this.Amount.ToString()
    }
}
$item = [FormattableItem_7]::new()
if ($item.ToString("HEX") -ne (700).ToString("X")) { Write-Host "FAIL: Formattable custom ToString failed"; exit 1 }
Write-Host "PASS"; exit 0
