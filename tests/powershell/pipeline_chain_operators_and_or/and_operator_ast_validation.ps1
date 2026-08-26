# vybe-test: powershell/pipeline_chain_operators_and_or/and_operator_ast_validation
$ast = [System.Management.Automation.Language.Parser]::ParseInput('
Write-Output 1 && Write-Output 2
', [ref]$null, [ref]$null)
$chainAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.PipelineChainAst] }, $true)
if ($chainAst -eq $null -or $chainAst.Operator.ToString() -ne "AndAnd") {
    Write-Host "FAIL: PipelineChainAst AST inspection failed"
    exit 1
}
Write-Host "PASS"
exit 0
