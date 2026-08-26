# vybe-test: powershell/classes_base_method_calls/base_method_cast_preserves_object_identity
class BaseIdent {
    [string]$Id = "ID001"
}
class SubIdent : BaseIdent {}
$sub = [SubIdent]::new()
$baseRef = [BaseIdent]$sub
if ($sub -ne $baseRef) {
    Write-Host "FAIL: Base cast reference identity failed"
    exit 1
}
Write-Host "PASS"
exit 0
