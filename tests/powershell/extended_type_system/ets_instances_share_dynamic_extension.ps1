# vybe-test: powershell/extended_type_system/ets_instances_share_dynamic_extension
# Dynamic properties added to a type are shared across all instances of that type
Update-TypeData -TypeName System.DateTime -MemberType ScriptProperty -MemberName YearString -Value { "YR-$($this.Year)" } -Force

$dt1 = [DateTime]::Parse("2024-01-01")
$dt2 = [DateTime]::Parse("2028-06-01")

if ($dt1.YearString -ne "YR-2024") {
    Write-Host "FAIL: instance 1 YearString mismatch, got: '$($dt1.YearString)'"
    exit 1
}

if ($dt2.YearString -ne "YR-2028") {
    Write-Host "FAIL: instance 2 YearString mismatch, got: '$($dt2.YearString)'"
    exit 1
}

Write-Host "PASS"
exit 0
