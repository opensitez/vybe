# vybe-test: powershell/extended_type_system/ets_scriptmethod_this_reference
# In an ETS ScriptMethod, $this refers to the current instance calling the method
Update-TypeData -TypeName System.String -MemberType ScriptMethod -MemberName BracketWrap -Value { "[$this]" } -Force

$wrapped = "antigravity".BracketWrap()

if ($wrapped -ne "[antigravity]") {
    Write-Host "FAIL: expected '[antigravity]', got: '$wrapped'"
    exit 1
}

Write-Host "PASS"
exit 0
