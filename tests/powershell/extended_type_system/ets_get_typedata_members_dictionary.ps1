# vybe-test: powershell/extended_type_system/ets_get_typedata_members_dictionary
# TypeData exposes a Members dictionary indexing extended properties and methods
Update-TypeData -TypeName System.DateTime -MemberType ScriptProperty -MemberName CustomMetaDay -Value { $this.Day } -Force
$td = Get-TypeData -TypeName System.DateTime

if ($null -eq $td.Members["CustomMetaDay"]) {
    Write-Host "FAIL: CustomMetaDay member not found in TypeData.Members dictionary"
    exit 1
}

Write-Host "PASS"
exit 0
