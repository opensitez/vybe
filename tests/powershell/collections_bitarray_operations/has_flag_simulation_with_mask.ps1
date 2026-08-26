# vybe-test: powershell/collections_bitarray_operations/has_flag_simulation_with_mask
$perms = [System.Collections.BitArray]::new(4, $false) # Read=0, Write=1, Exec=2
$perms.Set(0, $true) # Read
$perms.Set(2, $true) # Exec
$hasWrite = $perms.Get(1)
$hasExec = $perms.Get(2)
if ($hasWrite -or -not $hasExec) {
    Write-Host "FAIL: Permission flag simulation failed"
    exit 1
}
Write-Host "PASS"
exit 0
