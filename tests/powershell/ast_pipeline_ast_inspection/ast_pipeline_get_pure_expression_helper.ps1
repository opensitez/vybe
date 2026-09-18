# vybe-test: powershell/ast_pipeline_ast_inspection/ast_pipeline_get_pure_expression_helper
# PipelineAst.GetPureExpression() unpacks single-expression pipelines without redirections
$code = "`$counter + 1"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$pipeline = $ast.Find({ $args[0] -is [System.Management.Automation.Language.PipelineAst] }, $true)

$pureExpr = $pipeline.GetPureExpression()

if ($pureExpr -eq $null) {
    Write-Host "FAIL: GetPureExpression returned null for single pure expression"
    exit 1
}

if ($pureExpr -isnot [System.Management.Automation.Language.BinaryExpressionAst]) {
    Write-Host "FAIL: expected BinaryExpressionAst, got '$($pureExpr.GetType().Name)'"
    exit 1
}

# Conversely, a command pipeline returns null for GetPureExpression
$cmdCode = "Get-Process"
$cmdAst = [System.Management.Automation.Language.Parser]::ParseInput($cmdCode, [ref]$null, [ref]$null)
$cmdPipeline = $cmdAst.Find({ $args[0] -is [System.Management.Automation.Language.PipelineAst] }, $true)

if ($cmdPipeline.GetPureExpression() -ne $null) {
    Write-Host "FAIL: GetPureExpression was not null for command pipeline"
    exit 1
}

Write-Host "PASS"
exit 0
