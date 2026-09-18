# vybe-test: powershell/ast_error_handling_ast/ast_error_catch_single_type_constraint
# CatchClauseAst.CatchTypes contains 1 TypeConstraintAst matching the specified exception type
$code = "try { 1 } catch [System.ArgumentNullException] { 2 }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$catchAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.CatchClauseAst] }, $true)

if ($catchAst.CatchTypes.Count -ne 1) {
    Write-Host "FAIL: expected 1 catch type, got $($catchAst.CatchTypes.Count)"
    exit 1
}

$typeName = $catchAst.CatchTypes[0].TypeName.FullName

if ($typeName -ne "System.ArgumentNullException") {
    Write-Host "FAIL: catch type name mismatch: '$typeName'"
    exit 1
}

Write-Host "PASS"
exit 0
