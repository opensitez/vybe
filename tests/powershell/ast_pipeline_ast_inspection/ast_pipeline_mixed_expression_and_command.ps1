# vybe-test: powershell/ast_pipeline_ast_inspection/ast_pipeline_mixed_expression_and_command
# Piping an expression into a command produces CommandExpressionAst followed by CommandAst
$code = "1..10 | ForEach-Object { `$_ * 2 }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$pipeline = $ast.Find({ $args[0] -is [System.Management.Automation.Language.PipelineAst] }, $true)

$elem0 = $pipeline.PipelineElements[0]
$elem1 = $pipeline.PipelineElements[1]

if ($elem0 -isnot [System.Management.Automation.Language.CommandExpressionAst]) {
    Write-Host "FAIL: first element was not CommandExpressionAst, got '$($elem0.GetType().Name)'"
    exit 1
}

if ($elem1 -isnot [System.Management.Automation.Language.CommandAst]) {
    Write-Host "FAIL: second element was not CommandAst, got '$($elem1.GetType().Name)'"
    exit 1
}

Write-Host "PASS"
exit 0
