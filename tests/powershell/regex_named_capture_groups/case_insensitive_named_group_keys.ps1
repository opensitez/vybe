# vybe-test: powershell/regex_named_capture_groups/case_insensitive_named_group_keys
$str = "tag=prod"
$matched = $str -match "tag=(?<env>\w+)"
if ($Matches["ENV"] -ne "prod" -or $Matches["Env"] -ne "prod") {
    Write-Host "FAIL: `$Matches key case-insensitivity failed"
    exit 1
}
Write-Host "PASS"
exit 0
