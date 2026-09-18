# vybe-test: powershell/extended_type_system/ets_add_scriptmethod_with_parameters
# Update-TypeData adds a dynamic ScriptMethod accepting parameters
Update-TypeData -TypeName System.String -MemberType ScriptMethod -MemberName RepeatText -Value {
    param($count)
    $res = ""
    1..$count | ForEach-Object { $res += $this }
    $res
} -Force

$result = "ab".RepeatText(3)

if ($result -ne "ababab") {
    Write-Host "FAIL: expected 'ababab', got: '$result'"
    exit 1
}

Write-Host "PASS"
exit 0
