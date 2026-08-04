# vybe-test: powershell/comparison/case_insensitive_by_default
# PowerShell string comparisons are case-insensitive by default
if ("Hello" -ne "hello") { Write-Host "FAIL: case insensitive default"; exit 1 }
if ("WORLD" -ne "world") { Write-Host "FAIL: uppercase vs lowercase"; exit 1 }
# Explicit case-sensitive
if ("Hello" -ceq "hello") { Write-Host "FAIL: ceq should be case-sensitive"; exit 1 }
if ("Hello" -cne "Hello") { Write-Host "FAIL: ceq same string"; exit 1 }
Write-Host "PASS"
exit 0
