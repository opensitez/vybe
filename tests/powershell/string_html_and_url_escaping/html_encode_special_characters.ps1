# vybe-test: powershell/string_html_and_url_escaping/html_encode_special_characters
$raw = "<script>alert('xss'); & 'hello'</script>"
$enc = [System.Net.WebUtility]::HtmlEncode($raw)
if (-not $enc.Contains("&lt;script&gt;") -or -not $enc.Contains("&amp;")) { Write-Host "FAIL: HtmlEncode failed, got $enc"; exit 1 }
Write-Host "PASS"; exit 0
