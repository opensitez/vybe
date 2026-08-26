# vybe-test: powershell/classes_formattable_custom_tostring/formattable_custom_tostring_2
class FormattableItem_2 {
    [int]$Amount = 200
    [string]ToString([string]$format) {
        if ($format -eq "HEX") { return $this.Amount.ToString("X") }
        return $this.Amount.ToString()
    }
}
$item = [FormattableItem_2]::new()
if ($item.ToString("HEX") -ne (200).ToString("X")) { Write-Host "FAIL: Formattable custom ToString failed"; exit 1 }
Write-Host "PASS"; exit 0
