# vybe-test: powershell/out_string_cmdlet/out_string_stream_parameter_emits_array_of_strings
# The -Stream parameter causes Out-String to emit each line as an individual [string]
$lines = @("alpha", "beta", "gamma")
$streamed = @($lines | Out-String -Stream)

if ($streamed.Count -ne 3) {
    Write-Host "FAIL: expected 3 streamed string lines, got $($streamed.Count)"
    exit 1
}

if ($streamed[0] -ne "alpha" -or $streamed[1] -ne "beta" -or $streamed[2] -ne "gamma") {
    Write-Host "FAIL: streamed elements mismatch: @($($streamed -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
