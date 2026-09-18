# vybe-test: powershell/select_string_cmdlet/select_string_regex_named_groups
# MatchInfo.Matches exposes named capture groups via indexer lookup
$line = "client_ip=192.168.1.50 method=GET endpoint=/api/v1/data"
$match = $line | Select-String -Pattern "method=(?<verb>[A-Z]+)\s+endpoint=(?<path>\S+)"

if ($null -eq $match) {
    Write-Host "FAIL: pattern with named groups failed to match"
    exit 1
}

$verb = $match.Matches[0].Groups['verb'].Value
$path = $match.Matches[0].Groups['path'].Value

if ($verb -ne "GET" -or $path -ne "/api/v1/data") {
    Write-Host "FAIL: named capture mismatch, verb='$verb', path='$path'"
    exit 1
}

Write-Host "PASS"
exit 0
