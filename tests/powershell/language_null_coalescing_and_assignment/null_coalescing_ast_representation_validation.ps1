# vybe-test: powershell/language_null_coalescing_and_assignment/null_coalescing_ast_representation_validation
$ast = [System.Management.Automation.Language.Parser]::ParseInput('
$x = $a ?? $b
', [ref]$null, [ref]$null)
$nullCoalesceAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.BinaryExpressionAst] }, $true)
if ($nullCoalesceAst -eq $null) {
    Write-Host "FAIL: Null coalescing BinaryExpressionAst AST validation failed"
    exit 1
}
Write-Host "PASS"
exit 0
