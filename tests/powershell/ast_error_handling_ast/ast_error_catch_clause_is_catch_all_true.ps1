# vybe-test: powershell/ast_error_handling_ast/ast_error_catch_clause_is_catch_all_true
# An untyped catch block sets CatchClauseAst.IsCatchAll to true and CatchTypes is empty
$code = "try { 1 } catch { 2 }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$catchAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.CatchClauseAst] }, $true)

if (-not $catchAst.IsCatchAll) {
    Write-Host "FAIL: IsCatchAll was false for untyped catch block"
    exit 1
}

if ($catchAst.CatchTypes.Count -ne 0) {
    Write-Host "FAIL: CatchTypes was not empty for untyped catch block"
    exit 1
}

Write-Host "PASS"
exit 0
