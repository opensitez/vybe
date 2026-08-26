# vybe-test: powershell/regex_named_capture_groups/named_group_in_scriptblock_replace
$re = [regex]::new("(?<num>\d+)")
$res = $re.Replace("Items: 5 and 10", {
    param($match)
    return ([int]$match.Groups["num"].Value * 2).ToString()
})
if ($res -ne "Items: 10 and 20") {
    Write-Host "FAIL: Scriptblock replace with named group failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
