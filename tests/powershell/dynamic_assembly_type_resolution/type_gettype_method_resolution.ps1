# vybe-test: powershell/dynamic_assembly_type_resolution/type_gettype_method_resolution
$t = [System.Type]::GetType("System.Int32")
if ($t -ne [int]) {
    Write-Host "FAIL: Type.GetType('System.Int32') failed"
    exit 1
}
Write-Host "PASS"
exit 0
