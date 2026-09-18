# vybe-test: powershell/ast_error_handling_ast/ast_error_trap_typed_statement
# A typed trap statement sets TrapStatementAst.TrapType to the specified TypeConstraintAst
$code = "trap [System.DivideByZeroException] { 'Trapped zero divide' }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$trapAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.TrapStatementAst] }, $true)

if ($trapAst.TrapType -eq $null) {
    Write-Host "FAIL: TrapType was null for typed trap"
    exit 1
}

$typeName = $trapAst.TrapType.TypeName.FullName

if ($typeName -ne "System.DivideByZeroException") {
    Write-Host "FAIL: TrapType FullName mismatch: '$typeName'"
    exit 1
}

Write-Host "PASS"
exit 0
