# vybe-test: powershell/ast_loop_statements/ast_loop_foreach_nested_loop_hierarchy
# Nested foreach loops form an unambiguous ancestor-descendant hierarchy in the AST
$code = @"
foreach (`$row in `$matrix) {
    foreach (`$cell in `$row) {
        `$cell * 10
    }
}
"@

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$loops = @($ast.FindAll({ $args[0] -is [System.Management.Automation.Language.ForEachStatementAst] }, $true))

if ($loops.Count -ne 2) {
    Write-Host "FAIL: expected 2 foreach loops, got $($loops.Count)"
    exit 1
}

$outer = $loops[0]
$inner = $loops[1]

if ($outer.Variable.VariablePath.UserPath -ne "row") {
    Write-Host "FAIL: outer loop variable mismatch"
    exit 1
}

if ($inner.Variable.VariablePath.UserPath -ne "cell") {
    Write-Host "FAIL: inner loop variable mismatch"
    exit 1
}

Write-Host "PASS"
exit 0
