# vybe-test: powershell/ast_error_handling_ast/ast_error_throw_rethrow_pipeline_is_null
# Bare throw; statement (re-throw) has ThrowStatementAst.Pipeline equal to $null
$code = "try { 1 } catch { throw }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$throwAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.ThrowStatementAst] }, $true)

if ($throwAst.Pipeline -ne $null) {
    Write-Host "FAIL: rethrow Pipeline was not null"
    exit 1
}

Write-Host "PASS"
exit 0
