# vybe-test: powershell/classes_polymorphism_and_casting/polymorphic_collection_iteration
class Shape {
    [double]Area() { return 0.0 }
}
class Rectangle : Shape {
    [double]$W; [double]$H
    Rectangle([double]$w, [double]$h) { $this.W = $w; $this.H = $h }
    [double]Area() { return $this.W * $this.H }
}
class Square : Shape {
    [double]$S
    Square([double]$s) { $this.S = $s }
    [double]Area() { return $this.S * $this.S }
}
[Shape[]]$shapes = @([Rectangle]::new(3.0, 4.0), [Square]::new(5.0))
$sum = 0.0
foreach ($sh in $shapes) { $sum += $sh.Area() }
if ($sum -ne 37.0) { # 12 + 25 = 37
    Write-Host "FAIL: Polymorphic collection area sum expected 37, got $sum"
    exit 1
}
Write-Host "PASS"
exit 0
