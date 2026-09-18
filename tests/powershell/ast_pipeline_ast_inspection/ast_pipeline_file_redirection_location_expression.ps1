# vybe-test: powershell/ast_pipeline_ast_inspection/ast_pipeline_file_redirection_location_expression
# FileRedirectionAst.Location preserves the target path expression
$code = "Get-Process > '/tmp/custom_target_file.txt'"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$cmdAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true)
$redir = $cmdAst.Redirections[0]

$loc = $redir.Location

if ($loc -isnot [System.Management.Automation.Language.StringConstantExpressionAst]) {
    Write-Host "FAIL: location expression was not StringConstantExpressionAst"
    exit 1
}

if ($loc.Value -ne "/tmp/custom_target_file.txt") {
    Write-Host "FAIL: location value mismatch: '$($loc.Value)'"
    exit 1
}

Write-Host "PASS"
exit 0
