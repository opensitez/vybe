# vybe-test: powershell/cmdlets/new_object_dotnet
$sb = New-Object System.Text.StringBuilder
$sb.Append("Hello") | Out-Null
$sb.Append(", ")    | Out-Null
$sb.Append("World") | Out-Null
$result = $sb.ToString()
if ($result -ne "Hello, World") {
    Write-Host "FAIL: '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
