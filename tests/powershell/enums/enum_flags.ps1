# vybe-test: powershell/enums/enum_flags
[Flags()] enum Permission {
    None    = 0
    Read    = 1
    Write   = 2
    Execute = 4
}
$rw = [Permission]::Read -bor [Permission]::Write
if ([int]$rw -ne 3) { Write-Host "FAIL: Read|Write should be 3, got $([int]$rw)"; exit 1 }
# Test individual flag presence
$hasRead  = ($rw -band [Permission]::Read)    -ne 0
$hasExec  = ($rw -band [Permission]::Execute) -ne 0
if (-not $hasRead) { Write-Host "FAIL: should have Read";      exit 1 }
if ($hasExec)      { Write-Host "FAIL: should not have Exec";  exit 1 }
Write-Host "PASS"
exit 0
