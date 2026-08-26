# vybe-test: powershell/regex_named_capture_groups/dot_property_access_on_matches
$str = "server=db01;port=5432"
$matched = $str -match "server=(?<host>[^;]+);port=(?<port>\d+)"
if (-not $matched -or $Matches.host -ne "db01" -or $Matches.port -ne "5432") {
    Write-Host "FAIL: Dot property access on `$Matches failed"
    exit 1
}
Write-Host "PASS"
exit 0
