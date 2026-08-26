# vybe-test: powershell/string_encoding_utf8/decoder_stateful_stream_simulation
$enc = [System.Text.Encoding]::UTF8
$dec = $enc.GetDecoder()
[byte[]]$b1 = @(0xC3) # first half of é
[byte[]]$b2 = @(0xA9) # second half of é
[char[]]$c1 = New-Object char[] 1
[char[]]$c2 = New-Object char[] 1
$dec.GetChars($b1, 0, 1, $c1, 0)
$dec.GetChars($b2, 0, 1, $c2, 0)
if ($c2[0] -ne [char]'`u{00E9}') {
    Write-Host "FAIL: Stateful Decoder character decoding failed"
    exit 1
}
Write-Host "PASS"
exit 0
