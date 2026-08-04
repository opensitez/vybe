# vybe-test: powershell/strings/string_format_method
$result = [string]::Format("Hello, {0}! You are {1} years old.", "Bob", 25)
if ($result -ne "Hello, Bob! You are 25 years old.") {
    Write-Host "FAIL: expected 'Hello, Bob! You are 25 years old.', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
