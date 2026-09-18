# vybe-test: powershell/ast_loop_statements/ast_loop_continue_statement_labeled
# Labeled continue targetLoop parses Label as StringConstantExpressionAst with matching value
$code = ":targetLoop while (`$true) { continue targetLoop }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$contAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.ContinueStatementAst] }, $true)

if ($contAst.Label -eq $null) {
    Write-Host "FAIL: labeled continue had null Label"
    exit 1
}

if ($contAst.Label.Value -ne "targetLoop") {
    Write-Host "FAIL: continue label value mismatch: '$($contAst.Label.Value)'"
    exit 1
}

Write-Host "PASS"
exit 0
