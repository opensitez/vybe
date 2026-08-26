# vybe-test: powershell/pipeline_clean_block_lifecycle/clean_block_reflection_ast_validation
$ast = [System.Management.Automation.Language.Parser]::ParseInput('
function Test-AST {
    begin {}
    process {}
    end {}
    clean {}
}
', [ref]$null, [ref]$null)
$funcAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)
$body = $funcAst.Body
if ($body.CleanBlock -eq $null) {
    Write-Host "FAIL: CleanBlock AST representation check failed"
    exit 1
}
Write-Host "PASS"
exit 0
