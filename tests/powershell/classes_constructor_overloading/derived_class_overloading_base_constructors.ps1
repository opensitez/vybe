# vybe-test: powershell/classes_constructor_overloading/derived_class_overloading_base_constructors
class BaseWidget {
    [int]$Width
    BaseWidget([int]$w) { $this.Width = $w }
}
class ColorWidget : BaseWidget {
    [string]$Color
    ColorWidget([int]$w) : base($w) { $this.Color = "Black" }
    ColorWidget([int]$w, [string]$c) : base($w) { $this.Color = $c }
}
$w1 = [ColorWidget]::new(100)
$w2 = [ColorWidget]::new(200, "Red")
if ($w1.Color -ne "Black" -or $w2.Width -ne 200 -or $w2.Color -ne "Red") {
    Write-Host "FAIL: Multiple derived constructors calling base failed"
    exit 1
}
Write-Host "PASS"
exit 0
