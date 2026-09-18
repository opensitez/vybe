# vybe-test: powershell/clixml_serialization_engine/clixml_export_import_primitive_integers
# Integer primitive values (Int32, Int64, negative numbers) preserve numeric values through CliXml
$tmp = [System.IO.Path]::GetTempFileName()
$nums = @([int]42, [long]9876543210123, [int]-1024)

$nums | Export-Clixml -Path $tmp
$restored = Import-Clixml -Path $tmp
Remove-Item -Force $tmp

if ($restored[0] -ne 42) {
    Write-Host "FAIL: Int32 value mismatch, got: $($restored[0])"
    exit 1
}

if ($restored[1] -ne 9876543210123) {
    Write-Host "FAIL: Int64 value mismatch, got: $($restored[1])"
    exit 1
}

if ($restored[2] -ne -1024) {
    Write-Host "FAIL: negative integer mismatch, got: $($restored[2])"
    exit 1
}

Write-Host "PASS"
exit 0
