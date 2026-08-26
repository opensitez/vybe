# vybe-test: powershell/string_encoding_utf8/codepage_property_65001
$enc = [System.Text.Encoding]::UTF8
if ($enc.CodePage -ne 65001) {
    Write-Host "FAIL: CodePage expected 65001, got $($enc.CodePage)"
    exit 1
}
Write-Host "PASS"
exit 0
