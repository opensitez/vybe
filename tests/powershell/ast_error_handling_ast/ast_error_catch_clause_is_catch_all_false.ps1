# vybe-test: powershell/ast_error_handling_ast/ast_error_catch_clause_is_catch_all_false
# A typed catch block sets CatchClauseAst.IsCatchAll to false
$code = "try { 1 } catch [System.IO.IOException] { 2 }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$catchAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.CatchClauseAst] }, $true)

if ($catchAst.IsCatchAll) {
    Write-Host "FAIL: IsCatchAll was unexpectedly true for typed catch block"
    exit 1
}

Write-Host "PASS"
exit 0
