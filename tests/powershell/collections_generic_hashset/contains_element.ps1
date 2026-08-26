# vybe-test: powershell/collections_generic_hashset/contains_element
$set = [System.Collections.Generic.HashSet[string]]::new([string[]]@("red", "green", "blue"))
if (-not $set.Contains("green") -or $set.Contains("yellow")) {
    Write-Host "FAIL: HashSet Contains check failed"
    exit 1
}
Write-Host "PASS"
exit 0
