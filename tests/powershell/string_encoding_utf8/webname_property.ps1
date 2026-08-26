# vybe-test: powershell/string_encoding_utf8/webname_property
$enc = [System.Text.Encoding]::UTF8
if ($enc.WebName -ne "utf-8") {
    Write-Host "FAIL: WebName expected utf-8, got $($enc.WebName)"
    exit 1
}
Write-Host "PASS"
exit 0
