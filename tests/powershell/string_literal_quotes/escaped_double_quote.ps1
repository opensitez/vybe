# vybe-test: powershell/string_literal_quotes/escaped_double_quote
# The OTHER way to spell a literal double quote: double it. `""` inside a
# double-quoted string is one quote, mirroring `''` inside a single-quoted one.
$s = "She said ""Hi"" loudly"
if ($s -eq 'She said "Hi" loudly') { Write-Host 'PASS'; exit 0 }
Write-Host "FAIL: got [$s]"
exit 1
