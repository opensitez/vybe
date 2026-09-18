# vybe-test: powershell/new_object_cmdlet/new_object_stringwriter_writes_and_tostring
# New-Object System.IO.StringWriter captures written text and exposes it via ToString()
$sw = New-Object System.IO.StringWriter
$sw.Write("line one")
$sw.Write(" line two")

if ($sw.ToString() -ne "line one line two") {
    Write-Host "FAIL: StringWriter content mismatch: '$($sw.ToString())'"
    exit 1
}

Write-Host "PASS"
exit 0
