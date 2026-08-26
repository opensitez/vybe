# vybe-test: powershell/string_encoding_utf8/nobom_preamble_empty
$encNoBom = [System.Text.UTF8Encoding]::new($false)
$preamble = $encNoBom.GetPreamble()
if ($preamble.Length -ne 0) {
    Write-Host "FAIL: NoBOM UTF8 encoding should have empty preamble"
    exit 1
}
Write-Host "PASS"
exit 0
