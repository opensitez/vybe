# vybe-test: powershell/convertto_html_cmdlet/convertto_html_xml_character_entity_escaping
# ConvertTo-Html escapes HTML special characters into standard XML entities
$payload = [pscustomobject]@{ Escaped = "Tom & Jerry <cat > mouse `"cartoon`"" }
$html = $payload | ConvertTo-Html -Fragment
$text = $html -join " "

$expected = "Tom &amp; Jerry &lt;cat &gt; mouse &quot;cartoon&quot;"
if ($text -notmatch [regex]::Escape($expected)) {
    Write-Host "FAIL: entity escaping mismatch, got: $text"
    exit 1
}

Write-Host "PASS"
exit 0
