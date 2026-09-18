# vybe-test: powershell/ast_error_handling_ast/ast_error_trap_untyped_statement
# An untyped trap statement has TrapStatementAst.TrapType equal to $null
$code = "trap { continue }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$trapAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.TrapStatementAst] }, $true)

if ($trapAst.TrapType -ne $null) {
    Write-Host "FAIL: TrapType was not null for untyped trap"
    exit 1
}

Write-Host "PASS"
exit 0
