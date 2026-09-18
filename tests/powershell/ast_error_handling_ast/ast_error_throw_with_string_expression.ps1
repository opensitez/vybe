# vybe-test: powershell/ast_error_handling_ast/ast_error_throw_with_string_expression
# throw "Error string" parses as ThrowStatementAst with non-null Pipeline
$code = "throw 'Unrecoverable state encountered'"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$throwAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.ThrowStatementAst] }, $true)

if ($throwAst.Pipeline -eq $null) {
    Write-Host "FAIL: Pipeline was null on throw statement with payload"
    exit 1
}

$msg = $throwAst.Pipeline.GetPureExpression().Value
if ($msg -ne "Unrecoverable state encountered") {
    Write-Host "FAIL: thrown message mismatch: '$msg'"
    exit 1
}

Write-Host "PASS"
exit 0
