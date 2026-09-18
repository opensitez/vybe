# vybe-test: powershell/select_string_cmdlet/select_string_pipeline_streaming_lines
$logEntries = @(
    "2026-01-01 [INFO] Service started",
    "2026-01-01 [WARN] Slow database response",
    "2026-01-01 [ERROR] Failed to bind port 8080"
)

# Piping string collections streams each matching entry as a separate MatchInfo
$matches = @($logEntries | Select-String -Pattern "\[ERROR\]")

if ($matches.Count -ne 1) {
    Write-Host "FAIL: expected 1 error entry, got $($matches.Count)"
    exit 1
}

if ($matches[0].Line -ne "2026-01-01 [ERROR] Failed to bind port 8080") {
    Write-Host "FAIL: matched line content mismatch: '$($matches[0].Line)'"
    exit 1
}

Write-Host "PASS"
exit 0
