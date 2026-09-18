# vybe-test: powershell/new_object_cmdlet/new_object_stringreader_readline_sequence
# New-Object System.IO.StringReader reads lines from an in-memory string source
$sr = New-Object System.IO.StringReader "line1`nline2`nline3"

$lines = @()
while ($null -ne ($line = $sr.ReadLine())) {
    $lines += $line
}
$sr.Dispose()

if ($lines.Count -ne 3 -or $lines[1] -ne "line2") {
    Write-Host "FAIL: expected 3 lines with second='line2', got Count=$($lines.Count), [1]='$($lines[1])'"
    exit 1
}

Write-Host "PASS"
exit 0
