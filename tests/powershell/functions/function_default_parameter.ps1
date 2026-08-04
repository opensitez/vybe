# vybe-test: powershell/functions/function_default_parameter
function Greet {
    param($name = "World")
    return "Hello, $name"
}
$result = Greet
if ($result -ne "Hello, World") {
    Write-Host "FAIL: expected 'Hello, World', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
