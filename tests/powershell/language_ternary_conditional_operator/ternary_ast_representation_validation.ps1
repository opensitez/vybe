# vybe-test: powershell/language_ternary_conditional_operator/ternary_ast_representation_validation
$ast = [System.Management.Automation.Language.Parser]::ParseInput('
$x = $a ? $b : $c
', [ref]$null, [ref]$null)
$ternaryAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.TernaryExpressionAst] }, $true)
if ($ternaryAst -eq $null) {
    Write-Host "FAIL: TernaryExpressionAst AST validation failed"
    exit 1
}
Write-Host "PASS"
exit 0
