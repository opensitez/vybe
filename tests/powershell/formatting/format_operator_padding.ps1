# vybe-test: powershell/formatting/format_operator_padding
$left  = "{0,-10}" -f "hello"   # left-aligned, 10 chars
$right = "{0,10}"  -f "hello"   # right-aligned, 10 chars
if ($left  -ne "hello     ") { Write-Host "FAIL: left pad '$left'";  exit 1 }
if ($right -ne "     hello") { Write-Host "FAIL: right pad '$right'"; exit 1 }
Write-Host "PASS"
exit 0
