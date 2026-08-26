# vybe-test: powershell/json_nested_payload_depth/depth_with_custom_class_instance
class ChildNode { [string]$Val = "child" }
class ParentNode { [ChildNode]$Child = [ChildNode]::new() }
class GrandParentNode { [ParentNode]$Parent = [ParentNode]::new() }
$gp = [GrandParentNode]::new()
$json = $gp | ConvertTo-Json -Depth 4
$recovered = $json | ConvertFrom-Json
if ($recovered.Parent.Child.Val -ne "child") {
    Write-Host "FAIL: Depth with custom class hierarchy failed"
    exit 1
}
Write-Host "PASS"
exit 0
