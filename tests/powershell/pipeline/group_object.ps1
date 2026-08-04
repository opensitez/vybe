# vybe-test: powershell/pipeline/group_object
$data = @("apple","banana","avocado","blueberry","cherry")
$groups = $data | Group-Object { $_.Substring(0,1) }
$aGroup = $groups | Where-Object Name -eq "a"
if ($aGroup.Count -ne 2) {
    Write-Host "FAIL: expected 2 'a' words, got $($aGroup.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
