# vybe-test: powershell/extended_type_system/ets_overwrite_without_force_throws_exception
# Attempting to overwrite an existing ETS member without -Force throws an error
class OverwriteTarget { [string]$V = "Initial" }

Update-TypeData -TypeName OverwriteTarget -MemberType ScriptProperty -MemberName DynProp -Value { "First" } -Force

$threwError = $false
try {
    Update-TypeData -TypeName OverwriteTarget -MemberType ScriptProperty -MemberName DynProp -Value { "Second" } -ErrorAction Stop
} catch {
    $threwError = $true
}

if (-not $threwError) {
    Write-Host "FAIL: re-registering member without -Force did not throw"
    exit 1
}

Write-Host "PASS"
exit 0
