# vybe-test: powershell/ast_assignment_statement_ast/ast_assignment_scoped_variable_left
# Assignment to a scoped variable $global:config retains scope details on Left.VariablePath
$code = "`$global:SharedClusterState = 'Healthy'"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$assignAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true)

$vp = $assignAst.Left.VariablePath

if (-not $vp.IsGlobal) {
    Write-Host "FAIL: IsGlobal was false on assigned scoped variable"
    exit 1
}

if ($vp.UserPath -ne "global:SharedClusterState") {
    Write-Host "FAIL: UserPath mismatch: '$($vp.UserPath)'"
    exit 1
}

Write-Host "PASS"
exit 0
