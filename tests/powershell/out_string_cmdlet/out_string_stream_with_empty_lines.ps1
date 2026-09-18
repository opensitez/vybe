# vybe-test: powershell/out_string_cmdlet/out_string_stream_with_empty_lines
# Streaming input containing empty strings preserves empty lines as separate strings
$lines = @("header", "", "footer")
$streamed = @($lines | Out-String -Stream)

if ($streamed.Count -ne 3) {
    Write-Host "FAIL: expected 3 items in stream, got $($streamed.Count)"
    exit 1
}

if ($streamed[1] -ne "") {
    Write-Host "FAIL: interior blank line was not preserved, got: '$($streamed[1])'"
    exit 1
}

Write-Host "PASS"
exit 0
