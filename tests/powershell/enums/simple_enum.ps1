# vybe-test: powershell/enums/simple_enum
enum Color {
    Red = 1
    Green = 2
    Blue = 3
}

$color = [Color]::Green
if ($color -ne 2) {
    Write-Host "FAIL: expected 2, got $color"
    exit 1
}
Write-Host "PASS"
exit 0
