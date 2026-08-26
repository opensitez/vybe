# vybe-test: powershell/string_builder_operations/appendline_with_newline
$sb = [System.Text.StringBuilder]::new()
$null = $sb.AppendLine("Line 1")
$null = $sb.Append("Line 2")
$lines = $sb.ToString() -split "`r?`n"
if ($lines.Length -lt 2 -or $lines[0] -ne "Line 1" -or $lines[1] -ne "Line 2") {
    Write-Host "FAIL: AppendLine failed"
    exit 1
}
Write-Host "PASS"
exit 0
