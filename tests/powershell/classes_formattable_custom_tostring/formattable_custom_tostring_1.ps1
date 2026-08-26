# vybe-test: powershell/classes_formattable_custom_tostring/formattable_custom_tostring_1
class FormattableItem_1 {
    [int]$Amount = 100
    [string]ToString([string]$format) {
        if ($format -eq "HEX") { return $this.Amount.ToString("X") }
        return $this.Amount.ToString()
    }
}
$item = [FormattableItem_1]::new()
if ($item.ToString("HEX") -ne (100).ToString("X")) { Write-Host "FAIL: Formattable custom ToString failed"; exit 1 }
Write-Host "PASS"; exit 0
