# vybe-test: powershell/pipeline_object_capture/capture_in_while_condition
$queue = @(1, 2, 3)
$collected = @()
while ($item = $queue | Select-Object -First 1) {
    $collected += $item
    $queue = $queue | Select-Object -Skip 1
}
if ($collected.Count -ne 3 -or $collected[2] -ne 3) {
    Write-Host "FAIL: capture in while condition expected 3 items"
    exit 1
}
Write-Host "PASS"
exit 0
