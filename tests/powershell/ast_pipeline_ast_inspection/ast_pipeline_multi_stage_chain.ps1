# vybe-test: powershell/ast_pipeline_ast_inspection/ast_pipeline_multi_stage_chain
# Chaining three commands with pipes produces a PipelineAst with PipelineElements.Count equal to 3
$code = "Get-Service | Where-Object Status -eq 'Running' | Sort-Object DisplayName"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$pipeline = $ast.Find({ $args[0] -is [System.Management.Automation.Language.PipelineAst] }, $true)

if ($pipeline.PipelineElements.Count -ne 3) {
    Write-Host "FAIL: expected 3 pipeline elements, got $($pipeline.PipelineElements.Count)"
    exit 1
}

$names = @($pipeline.PipelineElements | ForEach-Object { $_.GetCommandName() })
if ($names[0] -ne "Get-Service" -or $names[1] -ne "Where-Object" -or $names[2] -ne "Sort-Object") {
    Write-Host "FAIL: pipeline command names mismatch: $($names -join ', ')"
    exit 1
}

Write-Host "PASS"
exit 0
