# vybe-test: powershell/type_char_classification_methods/char_arithmetic_offset
$ch = [char]'A'
$next = [char]([int]$ch + 1)
if ($next -ne [char]'B') {
    Write-Host "FAIL: Char arithmetic offset failed"
    exit 1
}
Write-Host "PASS"
exit 0
