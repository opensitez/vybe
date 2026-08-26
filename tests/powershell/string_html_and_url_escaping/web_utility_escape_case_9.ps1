# vybe-test: powershell/string_html_and_url_escaping/web_utility_escape_case_9
$raw = "param_9=value & data/test"
$enc = [System.Net.WebUtility]::UrlEncode($raw)
$dec = [System.Net.WebUtility]::UrlDecode($enc)
if ($dec -ne $raw) { Write-Host "FAIL: Url roundtrip failed"; exit 1 }
Write-Host "PASS"; exit 0
