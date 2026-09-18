# vybe-test: powershell/convertto_html_cmdlet/convertto_html_empty_collection_renders_table
# Piping an empty collection into ConvertTo-Html produces an empty table structure without throwing
$html = @() | ConvertTo-Html -Fragment

if ($null -eq $html) {
    Write-Host "FAIL: ConvertTo-Html with empty input returned `$null"
    exit 1
}

$text = $html -join " "
if ($text -notmatch "<table>.*</table>") {
    Write-Host "FAIL: expected table tag in empty collection HTML output, got: $text"
    exit 1
}

Write-Host "PASS"
exit 0
