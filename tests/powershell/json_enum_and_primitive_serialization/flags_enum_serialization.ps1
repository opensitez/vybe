# vybe-test: powershell/json_enum_and_primitive_serialization/flags_enum_serialization
[System.FlagsAttribute()]
enum Perms { Read = 1; Write = 2 }
$obj = @{ Perm = [Perms]::Read -bor [Perms]::Write }
$json = $obj | ConvertTo-Json
$recovered = $json | ConvertFrom-Json
if ($recovered.Perm -ne 3 -and $recovered.Perm -ne "Read, Write") {
    Write-Host "FAIL: Flags enum serialization failed"
    exit 1
}
Write-Host "PASS"
exit 0
