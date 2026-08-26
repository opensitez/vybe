# vybe-test: powershell/classes_this_chaining_and_scope/this_passing_to_child_object_as_parent_pointer
class ParentNode {
    [string]$Name = "Parent"
    [ChildNode]$Child
    ParentNode() {
        $this.Child = [ChildNode]::new($this)
    }
}
class ChildNode {
    [ParentNode]$Parent
    ChildNode([ParentNode]$p) { $this.Parent = $p }
    [string]GetParentName() { return $this.Parent.Name }
}
$pn = [ParentNode]::new()
if ($pn.Child.GetParentName() -ne "Parent") {
    Write-Host "FAIL: Parent-child `$this pointer linking failed"
    exit 1
}
Write-Host "PASS"
exit 0
