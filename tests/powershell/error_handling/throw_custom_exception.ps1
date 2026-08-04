# vybe-test: powershell/error_handling/throw_custom_exception
function Validate([int]$n) {
    if ($n -lt 0) { throw [System.ArgumentOutOfRangeException]::new("n", "Must be non-negative") }
    return $n * 2
}
$caught = $false
try {
    Validate -1
} catch [System.ArgumentOutOfRangeException] {
    $caught = $true
}
if (-not $caught) { Write-Host "FAIL: should have caught ArgumentOutOfRangeException"; exit 1 }
$ok = Validate 5
if ($ok -ne 10) { Write-Host "FAIL: Validate(5) should be 10"; exit 1 }
Write-Host "PASS"
exit 0
