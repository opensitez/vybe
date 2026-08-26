# vybe-test: powershell/classes_this_chaining_and_scope/this_in_recursive_tree_node
class TreeNode {
    [string]$Value
    [TreeNode]$Left
    [TreeNode]$Right
    TreeNode([string]$v) { $this.Value = $v }
    [string]GetInOrder() {
        $res = ""
        if ($this.Left -ne $null) { $res += $this.Left.GetInOrder() + "," }
        $res += $this.Value
        if ($this.Right -ne $null) { $res += "," + $this.Right.GetInOrder() }
        return $res
    }
}
$root = [TreeNode]::new("B")
$root.Left = [TreeNode]::new("A")
$root.Right = [TreeNode]::new("C")
$order = $root.GetInOrder()
if ($order -ne "A,B,C") {
    Write-Host "FAIL: Recursive tree node `$this failed, got '$order'"
    exit 1
}
Write-Host "PASS"
exit 0
