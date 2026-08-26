# vybe-test: powershell/regex_lookaround_assertions/lookaround_stripping_html_tags_content_only
$html = "<b>Bold</b> and <i>Italic</i>"
$re = [regex]::new("(?<=>)[^<]+(?=<)")
$matches = @($re.Matches($html) | ForEach-Object { $_.Value.Trim() } | Where-Object { $_ -ne "" })
if ($matches.Count -ne 3 -or $matches[0] -ne "Bold" -or $matches[2] -ne "Italic") {
    Write-Host "FAIL: HTML tag content extraction via lookaround failed"
    exit 1
}
Write-Host "PASS"
exit 0
