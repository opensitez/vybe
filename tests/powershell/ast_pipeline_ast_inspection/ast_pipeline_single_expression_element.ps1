# vybe-test: powershell/ast_pipeline_ast_inspection/ast_pipeline_single_expression_element
# An isolated expression statement is represented by a PipelineAst containing one CommandExpressionAst element
$code = "40 + 2"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$pipeline = $ast.Find({ $args[0] -is [System.Management.Automation.Language.PipelineAst] }, $true)

if ($pipeline.PipelineElements.Count -ne 1) {
    Write-Host "FAIL: expected 1 pipeline element, got $($pipeline.PipelineElements.Count)"
    exit 1
}

$firstElem = $pipeline.PipelineElements[0]
if ($firstElem -isnot [System.Management.Automation.Language.CommandExpressionAst]) {
    Write-Host "FAIL: expected CommandExpressionAst element, got '$($firstElem.GetType().Name)'"
    exit 1
}

if ($firstElem.Expression -isnot [System.Management.Automation.Language.BinaryExpressionAst]) {
    Write-Host "FAIL: inner expression was not BinaryExpressionAst"
    exit 1
}

Write-Host "PASS"
exit 0
