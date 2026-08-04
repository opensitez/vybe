# vybe-test: powershell/classes/class_constructor_overload
class Rectangle {
    [double]$Width
    [double]$Height
    Rectangle([double]$w, [double]$h) { $this.Width = $w; $this.Height = $h }
    Rectangle([double]$side) { $this.Width = $side; $this.Height = $side }
    [double]Area() { return $this.Width * $this.Height }
}
$rect = [Rectangle]::new(4, 5)
$sq   = [Rectangle]::new(6)
if ($rect.Area() -ne 20) { Write-Host "FAIL: rect area"; exit 1 }
if ($sq.Area()   -ne 36) { Write-Host "FAIL: square area"; exit 1 }
Write-Host "PASS"
exit 0
