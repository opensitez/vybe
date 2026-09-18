# vybe-test: powershell/ast_loop_statements/ast_loop_break_statement_labeled
# Labeled break statement break outerLoop parses Label as StringConstantExpressionAst with matching value
$code = ":outerLoop while (`$true) { break outerLoop }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$breakAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.BreakStatementAst] }, $true)

if ($breakAst.Label -eq $null) {
    Write-Host "FAIL: labeled break had null Label"
    exit 1
}

if ($breakAst.Label.Value -ne "outerLoop") {
    Write-Host "FAIL: break label value mismatch: '$($breakAst.Label.Value)'"
    exit 1
}

Write-Host "PASS"
exit 0
