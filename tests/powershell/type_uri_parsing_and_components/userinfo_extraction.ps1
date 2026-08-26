# vybe-test: powershell/type_uri_parsing_and_components/userinfo_extraction
$u = [uri]"https://alice:secret@example.com/repo"
if ($u.UserInfo -ne "alice:secret") {
    Write-Host "FAIL: UserInfo extraction failed"
    exit 1
}
Write-Host "PASS"
exit 0
