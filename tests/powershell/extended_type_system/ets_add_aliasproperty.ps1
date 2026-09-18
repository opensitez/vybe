# vybe-test: powershell/extended_type_system/ets_add_aliasproperty
# Update-TypeData adds an AliasProperty that forwards access to an existing property
Update-TypeData -TypeName System.DateTime -MemberType AliasProperty -MemberName Yr -Value "Year" -Force
$dt = [DateTime]::Parse("2026-08-15")

if ($dt.Yr -ne $dt.Year -or $dt.Yr -ne 2026) {
    Write-Host "FAIL: AliasProperty 'Yr' did not match 'Year', got: $($dt.Yr)"
    exit 1
}

Write-Host "PASS"
exit 0
