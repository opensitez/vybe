# vybe-test: powershell/dynamic_method_invocations_by_string/dynamic_method_non_existent_method_throws
$str = "text"
$badMethod = "NonExistentMethod"
$caught = $false
try {
    $x = $str.$badMethod()
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Calling non-existent dynamic method must throw"
    exit 1
}
Write-Host "PASS"
exit 0
