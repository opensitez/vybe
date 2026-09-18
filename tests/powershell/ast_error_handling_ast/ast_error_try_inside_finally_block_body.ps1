# vybe-test: powershell/ast_error_handling_ast/ast_error_try_inside_finally_block_body
# A TryStatementAst nested inside a finally block is properly located in the AST hierarchy
$code = @"
try {
    Write-Host "work"
} finally {
    try {
        Close-Resource
    } catch {
        Write-Warning "cleanup failed"
    }
}
"@

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$outerTry = $ast.Find({ $args[0] -is [System.Management.Automation.Language.TryStatementAst] }, $false)
$finallyBlock = $outerTry.Finally

$cleanupTry = $finallyBlock.Find({ $args[0] -is [System.Management.Automation.Language.TryStatementAst] }, $true)

if ($cleanupTry -eq $null) {
    Write-Host "FAIL: cleanup TryStatementAst not found inside finally block"
    exit 1
}

Write-Host "PASS"
exit 0
