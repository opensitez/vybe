# vybe-test: powershell/classes_formattable_custom_tostring/formattable_custom_tostring_19
class FormattableItem_19 {
    [int]$Amount = 1900
    [string]ToString([string]$format) {
        if ($format -eq "HEX") { return $this.Amount.ToString("X") }
        return $this.Amount.ToString()
    }
}
$item = [FormattableItem_19]::new()
if ($item.ToString("HEX") -ne (1900).ToString("X")) { Write-Host "FAIL: Formattable custom ToString failed"; exit 1 }
Write-Host "PASS"; exit 0
