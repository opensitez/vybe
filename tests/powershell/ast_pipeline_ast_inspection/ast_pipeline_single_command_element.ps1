# vybe-test: powershell/ast_pipeline_ast_inspection/ast_pipeline_single_command_element
# A single command call is represented by a PipelineAst containing one CommandAst element
$code = "Get-Process -Name 'pwsh'"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$pipeline = $ast.Find({ $args[0] -is [System.Management.Automation.Language.PipelineAst] }, $true)

if ($pipeline.PipelineElements.Count -ne 1) {
    Write-Host "FAIL: expected 1 pipeline element, got $($pipeline.PipelineElements.Count)"
    exit 1
}

$firstElem = $pipeline.PipelineElements[0]
if ($firstElem -isnot [System.Management.Automation.Language.CommandAst]) {
    Write-Host "FAIL: expected CommandAst element, got '$($firstElem.GetType().Name)'"
    exit 1
}

Write-Host "PASS"
exit 0
