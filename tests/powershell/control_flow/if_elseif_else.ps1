# vybe-test: powershell/control_flow/if_elseif_else
$x = 5
$result = ""
if ($x -eq 3) {
    $result = "three"
} elseif ($x -eq 5) {
    $result = "five"
} else {
    $result = "other"
}
if ($result -ne "five") {
    Write-Host "FAIL: expected 'five', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
