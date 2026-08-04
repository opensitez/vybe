# vybe-test: powershell/functions/return_early
function Test-Return {
    param($x)
    if ($x -lt 0) {
        return "negative"
    }
    return "positive"
}
$result = Test-Return -x -5
if ($result -ne "negative") {
    Write-Host "FAIL: expected 'negative', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
