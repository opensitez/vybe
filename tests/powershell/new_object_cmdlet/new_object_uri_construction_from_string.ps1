# vybe-test: powershell/new_object_cmdlet/new_object_uri_construction_from_string
# New-Object System.Uri -ArgumentList <url string> parses the URL and exposes Host
$uri = New-Object System.Uri -ArgumentList "https://example.com/path?q=1"

if ($uri.Host -ne "example.com") {
    Write-Host "FAIL: expected host 'example.com', got '$($uri.Host)'"
    exit 1
}

if ($uri.Scheme -ne "https") {
    Write-Host "FAIL: expected scheme 'https', got '$($uri.Scheme)'"
    exit 1
}

Write-Host "PASS"
exit 0
