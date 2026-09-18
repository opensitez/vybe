# vybe-test: powershell/ast_extent_source_mapping/ast_extent_pipeline_chain_operator_extents
# PipelineAst.Extent encompasses all commands chained together with pipeline pipes '|'
$code = "1..5 | ForEach-Object { `$_ * 2 } | Measure-Object -Sum"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$pipeAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.PipelineAst] }, $true)
$pipeText = $pipeAst.Extent.Text

if ($pipeText -ne $code) {
    Write-Host "FAIL: pipeline extent mismatch, expected '$code', got '$pipeText'"
    exit 1
}

# The pipeline contains 3 distinct pipeline elements
if ($pipeAst.PipelineElements.Count -ne 3) {
    Write-Host "FAIL: expected 3 pipeline elements, got $($pipeAst.PipelineElements.Count)"
    exit 1
}

Write-Host "PASS"
exit 0
