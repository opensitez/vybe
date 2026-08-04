# vybe-test: powershell/strings/string_isnullorempty
$empty = ""
$result = [string]::IsNullOrEmpty($empty)
if ($result -ne $true) {
    Write-Host "FAIL: expected True for empty string"
    exit 1
}
$nonempty = "text"
$result2 = [string]::IsNullOrEmpty($nonempty)
if ($result2 -ne $false) {
    Write-Host "FAIL: expected False for non-empty string"
    exit 1
}
Write-Host "PASS"
exit 0
