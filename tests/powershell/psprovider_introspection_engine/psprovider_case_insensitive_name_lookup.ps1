# vybe-test: powershell/psprovider_introspection_engine/psprovider_case_insensitive_name_lookup
# Provider name lookups in Get-PSProvider are case-insensitive
$lower = Get-PSProvider filesystem
$upper = Get-PSProvider FILESYSTEM

if ($lower.Name -ne "FileSystem" -or $upper.Name -ne "FileSystem") {
    Write-Host "FAIL: case-insensitive provider lookup failed: lower='$($lower.Name)', upper='$($upper.Name)'"
    exit 1
}

Write-Host "PASS"
exit 0
