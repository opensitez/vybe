# vybe-test: powershell/string_html_and_url_escaping/html_decode_roundtrip
$encoded = "&lt;div class=&quot;box&quot;&gt;Hello &amp; Welcome&lt;/div&gt;"
$dec = [System.Net.WebUtility]::HtmlDecode($encoded)
if ($dec -ne '<div class="box">Hello & Welcome</div>') { Write-Host "FAIL: HtmlDecode failed, got $dec"; exit 1 }
Write-Host "PASS"; exit 0
