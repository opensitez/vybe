# vybe-test: powershell/ast_parsing/ast_parsing_pipeline_ast
$sb = { 1..5 | ForEach-Object { $_ } }
$pipeAst = $sb.Ast.Find({ param($ast) $ast -is [System.Management.Automation.Language.PipelineAst] }, $true)
if ($pipeAst.PipelineElements.Count -ne 2) {
    Write-Host "FAIL: PipelineAst elements expected 2, got $($pipeAst.PipelineElements.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
