# vybe-test: powershell/ast_error_handling_ast/ast_error_try_without_catch_with_finally
# A try block without catch but with finally parses CatchClauses as empty and Finally as non-null
$code = "try { `$res = 100 } finally { `$res = 0 }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$tryAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.TryStatementAst] }, $true)

if ($tryAst.CatchClauses.Count -ne 0) {
    Write-Host "FAIL: CatchClauses was not empty in try-finally construct"
    exit 1
}

if ($tryAst.Finally -eq $null) {
    Write-Host "FAIL: Finally was null in try-finally construct"
    exit 1
}

Write-Host "PASS"
exit 0
