# vybe-test: powershell/classes_property_attributes/validaterange_on_int_property
class ScoreEntry {
    [ValidateRange(0, 100)][int]$Score
}
$s = [ScoreEntry]::new()
$s.Score = 85
$caught = $false
try {
    $s.Score = 150
} catch {
    $caught = $true
}
if ($s.Score -ne 85 -or -not $caught) {
    Write-Host "FAIL: ValidateRange on property failed"
    exit 1
}
Write-Host "PASS"
exit 0
