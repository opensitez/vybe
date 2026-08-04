# vybe-test: powershell/transcript/transcript_error_handling
try { Start-Transcript -Path 'invalid:\path' -ErrorAction Stop } catch { $caught = $true }
if (-not $caught) {
    Write-Host "FAIL: expected transcript error"
    exit 1
}
Write-Host 'PASS'
exit 0
