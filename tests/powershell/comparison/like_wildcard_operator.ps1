# vybe-test: powershell/comparison/like_wildcard_operator
if ("PowerShell" -notlike "Power*")   { Write-Host "FAIL: Power*";   exit 1 }
if ("PowerShell" -notlike "*Shell")   { Write-Host "FAIL: *Shell";   exit 1 }
if ("PowerShell" -notlike "*wer*")    { Write-Host "FAIL: *wer*";    exit 1 }
if ("PowerShell" -like "Python*")     { Write-Host "FAIL: Python*";  exit 1 }
Write-Host "PASS"
exit 0
