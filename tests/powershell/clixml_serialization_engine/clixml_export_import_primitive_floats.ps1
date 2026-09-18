# vybe-test: powershell/clixml_serialization_engine/clixml_export_import_primitive_floats
# Floating point numbers (<Db> tags) retain precision across CliXml round-tripping
$tmp = [System.IO.Path]::GetTempFileName()
$floats = @([double]3.1415926535, [double]-0.000123)

$floats | Export-Clixml -Path $tmp
$restored = Import-Clixml -Path $tmp
Remove-Item -Force $tmp

if ([Math]::Abs($restored[0] - 3.1415926535) -gt 0.000001) {
    Write-Host "FAIL: float precision lost, got: $($restored[0])"
    exit 1
}

if ([Math]::Abs($restored[1] - (-0.000123)) -gt 0.0000001) {
    Write-Host "FAIL: negative float precision lost, got: $($restored[1])"
    exit 1
}

Write-Host "PASS"
exit 0
