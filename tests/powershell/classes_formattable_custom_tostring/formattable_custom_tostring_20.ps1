# vybe-test: powershell/classes_formattable_custom_tostring/formattable_custom_tostring_20
class FormattableItem_20 {
    [int]$Amount = 2000
    [string]ToString([string]$format) {
        if ($format -eq "HEX") { return $this.Amount.ToString("X") }
        return $this.Amount.ToString()
    }
}
$item = [FormattableItem_20]::new()
if ($item.ToString("HEX") -ne (2000).ToString("X")) { Write-Host "FAIL: Formattable custom ToString failed"; exit 1 }
Write-Host "PASS"; exit 0
