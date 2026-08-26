# vybe-test: powershell/string_builder_operations/append_boolean_and_numeric_types
$sb = [System.Text.StringBuilder]::new()
$null = $sb.Append($true)
$null = $sb.Append(123)
$null = $sb.Append(4.5)
if ($sb.ToString() -ne "True1234.5") {
    Write-Host "FAIL: Append typed primitives failed, got $($sb.ToString())"
    exit 1
}
Write-Host "PASS"
exit 0
