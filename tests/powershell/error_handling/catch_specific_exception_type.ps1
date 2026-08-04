# vybe-test: powershell/error_handling/catch_specific_exception_type
$caught = ""
try {
    [int]::Parse("not-a-number")
} catch [System.FormatException] {
    $caught = "format"
} catch {
    $caught = "other"
}
if ($caught -ne "format") {
    Write-Host "FAIL: expected 'format', got '$caught'"
    exit 1
}
Write-Host "PASS"
exit 0
