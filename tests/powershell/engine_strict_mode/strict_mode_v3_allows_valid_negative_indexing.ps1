# vybe-test: powershell/engine_strict_mode/strict_mode_v3_allows_valid_negative_indexing
Set-StrictMode -Version 3.0

$arr = @(100, 200, 300, 400)

# StrictMode v3.0 must allow in-bounds negative indexes to resolve from the end of the collection
$caught = $false
$last = $null
$secondLast = $null
try {
    $last = $arr[-1]
    $secondLast = $arr[-2]
} catch {
    $caught = $true
}

if ($caught) {
    Write-Host "FAIL: in-bounds negative indexing threw under v3.0"
    exit 1
}

if ($last -ne 400 -or $secondLast -ne 300) {
    Write-Host "FAIL: expected last=400, secondLast=300, got last=$last, secondLast=$secondLast"
    exit 1
}

Write-Host "PASS"
exit 0
