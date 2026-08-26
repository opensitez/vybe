# vybe-test: powershell/classes_base_method_calls/base_constructor_call_with_base_keyword
class ParentResource {
    [string]$Id
    ParentResource([string]$id) { $this.Id = $id }
}
class ChildResource : ParentResource {
    [string]$Tag
    ChildResource([string]$id, [string]$tag) : base($id) {
        $this.Tag = $tag
    }
}
$cr = [ChildResource]::new("RES01", "Prod")
if ($cr.Id -ne "RES01" -or $cr.Tag -ne "Prod") {
    Write-Host "FAIL: Base keyword constructor invocation failed"
    exit 1
}
Write-Host "PASS"
exit 0
