# vybe-test: powershell/string_html_and_url_escaping/url_encode_query_parameters
$param = "hello world & foo=bar/baz?"
$enc = [System.Net.WebUtility]::UrlEncode($param)
if ($enc -ne "hello+world+&foo%3Dbar%2Fbaz%3F" -and -not $enc.Contains("%26") -and -not $enc.Contains("+")) {
    Write-Host "FAIL: UrlEncode query parameter failed, got $enc"; exit 1
}
Write-Host "PASS"; exit 0
