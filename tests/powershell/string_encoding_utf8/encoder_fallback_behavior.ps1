# vybe-test: powershell/string_encoding_utf8/encoder_fallback_behavior
$enc = [System.Text.UTF8Encoding]::new($false, $true) # throw on invalid
if ($enc.EncoderFallback -eq $null -or $enc.DecoderFallback -eq $null) {
    Write-Host "FAIL: Fallback properties check failed"
    exit 1
}
Write-Host "PASS"
exit 0
