# vybe-test: powershell/collections_generic_hashset/case_insensitive_string_hashset
$comp = [System.StringComparer]::OrdinalIgnoreCase
$set = [System.Collections.Generic.HashSet[string]]::new($comp)
$set.Add("HELLO")
if (-not $set.Contains("hello") -or $set.Add("Hello")) {
    Write-Host "FAIL: Case-insensitive HashSet failed"
    exit 1
}
Write-Host "PASS"
exit 0
