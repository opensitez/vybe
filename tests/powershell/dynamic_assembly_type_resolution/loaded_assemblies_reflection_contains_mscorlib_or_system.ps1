# vybe-test: powershell/dynamic_assembly_type_resolution/loaded_assemblies_reflection_contains_mscorlib_or_system
$assemblies = @([System.AppDomain]::CurrentDomain.GetAssemblies() | ForEach-Object { $_.GetName().Name })
if (-not ($assemblies -contains "System.Private.CoreLib" -or $assemblies -contains "mscorlib" -or $assemblies -contains "System.Runtime")) {
    Write-Host "FAIL: Loaded assemblies check failed"
    exit 1
}
Write-Host "PASS"
exit 0
