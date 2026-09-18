# vybe-test: powershell/new_object_cmdlet/new_object_stringbuilder_append_and_tostring
# New-Object creates a System.Text.StringBuilder; Append() adds text and ToString() retrieves it
$sb = New-Object System.Text.StringBuilder
$sb.Append("hello") | Out-Null
$sb.Append(" world") | Out-Null

if ($sb.ToString() -ne "hello world") {
    Write-Host "FAIL: StringBuilder result mismatch: '$($sb.ToString())'"
    exit 1
}

Write-Host "PASS"
exit 0
