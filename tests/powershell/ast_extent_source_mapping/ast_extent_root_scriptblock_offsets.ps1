# vybe-test: powershell/ast_extent_source_mapping/ast_extent_root_scriptblock_offsets
# ScriptBlockAst.Extent starts at offset 0 and spans the full length of the input script text
$code = "Write-Host 'Source mapping test'"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$extent = $ast.Extent

if ($extent.StartOffset -ne 0) {
    Write-Host "FAIL: root extent StartOffset was not 0, got $($extent.StartOffset)"
    exit 1
}

if ($extent.EndOffset -ne $code.Length) {
    Write-Host "FAIL: root extent EndOffset mismatch, expected $($code.Length), got $($extent.EndOffset)"
    exit 1
}

if ($extent.Text -ne $code) {
    Write-Host "FAIL: root extent Text mismatch"
    exit 1
}

Write-Host "PASS"
exit 0
