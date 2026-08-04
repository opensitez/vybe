# vybe-test: powershell/regex/regex_select_string
$lines = @("apple 1", "banana 2", "cherry 3")
$matched = $lines | Select-String -Pattern "\d+"
$count = ($matched | Measure-Object).Count
if ($count -ne 3) {
    Write-Host "FAIL: expected 3 matches, got $count"
    exit 1
}
Write-Host "PASS"
exit 0
