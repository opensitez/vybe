# vybe-test: powershell/scriptblocks/scriptblock_as_argument
function Apply {
    param([scriptblock]$Action, [int]$Value)
    return & $Action $Value
}
$double = { param($x) $x * 2 }
$result = Apply -Action $double -Value 15
if ($result -ne 30) {
    Write-Host "FAIL: expected 30, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
