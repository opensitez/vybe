# vybe-test: powershell/string_html_and_url_escaping/url_decode_roundtrip
$encoded = "name=John+Doe&city=New%20York"
$dec = [System.Net.WebUtility]::UrlDecode($encoded)
if ($dec -ne "name=John Doe&city=New York") { Write-Host "FAIL: UrlDecode failed, got $dec"; exit 1 }
Write-Host "PASS"; exit 0
