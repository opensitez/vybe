# vybe-test: powershell/classes_hidden_members/hidden_array_property
class ContainerList {
    hidden [int[]]$Items = @(10, 20, 30)
    [int]GetCount() { return $this.Items.Length }
}
$cl = [ContainerList]::new()
if ($cl.GetCount() -ne 3) {
    Write-Host "FAIL: Hidden array property failed"
    exit 1
}
Write-Host "PASS"
exit 0
