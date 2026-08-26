# vybe-test: powershell/classes_this_chaining_and_scope/this_type_name_inspection
class NameInspector {
    [string]GetTypeName() { return $this.GetType().Name }
}
$ni = [NameInspector]::new()
if ($ni.GetTypeName() -ne "NameInspector") {
    Write-Host "FAIL: `$this.GetType().Name inspection failed"
    exit 1
}
Write-Host "PASS"
exit 0
