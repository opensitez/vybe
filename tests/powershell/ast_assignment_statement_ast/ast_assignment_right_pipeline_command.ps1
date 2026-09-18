# vybe-test: powershell/ast_assignment_statement_ast/ast_assignment_right_pipeline_command
# Assigning the result of a command $procs = Get-Process sets Right to a PipelineAst
$code = "`$procs = Get-Process -Name 'pwsh'"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$assignAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true)

$right = $assignAst.Right

if ($right -isnot [System.Management.Automation.Language.PipelineAst]) {
    Write-Host "FAIL: Right was not PipelineAst, got '$($right.GetType().Name)'"
    exit 1
}

$cmd = $right.PipelineElements[0]
if ($cmd.GetCommandName() -ne "Get-Process") {
    Write-Host "FAIL: command name mismatch: '$($cmd.GetCommandName())'"
    exit 1
}

Write-Host "PASS"
exit 0
