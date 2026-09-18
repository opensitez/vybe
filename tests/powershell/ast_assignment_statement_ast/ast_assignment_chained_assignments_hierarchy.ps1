# vybe-test: powershell/ast_assignment_statement_ast/ast_assignment_chained_assignments_hierarchy
# Chained assignment $a = $b = $c = 0 produces 3 nested AssignmentStatementAst nodes
$code = "`$a = `$b = `$c = 0"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$assigns = @($ast.FindAll({ $args[0] -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true))

if ($assigns.Count -ne 3) {
    Write-Host "FAIL: expected 3 assignments in chain, got $($assigns.Count)"
    exit 1
}

$names = @($assigns | ForEach-Object { $_.Left.VariablePath.UserPath })

if ($names[0] -ne "a" -or $names[1] -ne "b" -or $names[2] -ne "c") {
    Write-Host "FAIL: chain variable names mismatch: $($names -join ', ')"
    exit 1
}

Write-Host "PASS"
exit 0
