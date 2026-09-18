# vybe-test: powershell/clixml_serialization_engine/clixml_pipeline_streaming_multiple_objects
# Streaming multiple individual objects through Export-Clixml serializes all of them
$tmp = [System.IO.Path]::GetTempFileName()
$items = @("alpha_item", "beta_item", "gamma_item")

$items | Export-Clixml -Path $tmp
$restored = @(Import-Clixml -Path $tmp)
Remove-Item -Force $tmp

if ($restored.Count -ne 3) {
    Write-Host "FAIL: expected 3 items restored, got $($restored.Count)"
    exit 1
}

if ($restored[0] -ne "alpha_item" -or $restored[1] -ne "beta_item" -or $restored[2] -ne "gamma_item") {
    Write-Host "FAIL: stream sequence mismatch: @($($restored -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
