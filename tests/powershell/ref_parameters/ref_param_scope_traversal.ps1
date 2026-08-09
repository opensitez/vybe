# vybe-test: powershell/ref_parameters/ref_param_scope_traversal
$script:globalVal = 1
function Outer-Scope {
    function Inner-Mutate([ref]$r) {
        $r.Value = 99
    }
    Inner-Mutate ([ref]$script:globalVal)
}
Outer-Scope
if ($script:globalVal -ne 99) {
    Write-Host "FAIL: scope-traversing [ref] expected 99, got $script:globalVal"
    exit 1
}
Write-Host "PASS"
exit 0
