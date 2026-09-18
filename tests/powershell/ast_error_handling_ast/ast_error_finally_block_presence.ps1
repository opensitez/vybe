# vybe-test: powershell/ast_error_handling_ast/ast_error_finally_block_presence
# TryStatementAst.Finally is a non-null StatementBlockAst when finally block is provided
$code = "try { 1 } finally { `$cleanup = `$true }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$tryAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.TryStatementAst] }, $true)

if ($tryAst.Finally -eq $null) {
    Write-Host "FAIL: Finally was null when finally block was present"
    exit 1
}

if ($tryAst.Finally -isnot [System.Management.Automation.Language.StatementBlockAst]) {
    Write-Host "FAIL: expected Finally to be StatementBlockAst, got '$($tryAst.Finally.GetType().Name)'"
    exit 1
}

Write-Host "PASS"
exit 0
