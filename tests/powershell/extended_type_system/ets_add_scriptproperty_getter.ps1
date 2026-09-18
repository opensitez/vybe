# vybe-test: powershell/extended_type_system/ets_add_scriptproperty_getter
# Update-TypeData adds a dynamic ScriptProperty getter to a .NET type
Update-TypeData -TypeName System.DateTime -MemberType ScriptProperty -MemberName CustomDayName -Value { $this.ToString("dddd") } -Force
$dt = [DateTime]::Parse("2026-01-01") # Thursday

if ($dt.CustomDayName -ne "Thursday") {
    Write-Host "FAIL: expected 'Thursday', got: '$($dt.CustomDayName)'"
    exit 1
}

Write-Host "PASS"
exit 0
