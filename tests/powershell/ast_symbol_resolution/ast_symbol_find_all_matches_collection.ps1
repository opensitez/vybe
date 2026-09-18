# vybe-test: powershell/ast_symbol_resolution/ast_symbol_find_all_matches_collection
# Ast.FindAll returns all nodes matching the predicate in source order
$code = @"
`$a = 1
`$b = 2
`$c = 3
"@

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$vars = @($ast.FindAll({ $args[0] -is [System.Management.Automation.Language.VariableExpressionAst] }, $true))

if ($vars.Count -ne 3) {
    Write-Host "FAIL: expected 3 variable ASTs, got $($vars.Count)"
    exit 1
}

$expectedNames = @("a", "b", "c")
for ($i = 0; $i -lt 3; $i++) {
    if ($vars[$i].VariablePath.UserPath -ne $expectedNames[$i]) {
        Write-Host "FAIL: at index ${i}, expected '$($expectedNames[$i])', got '$($vars[$i].VariablePath.UserPath)'"
        exit 1
    }
}

Write-Host "PASS"
exit 0
