# vybe-test: powershell/ast_switch_statement_ast/ast_switch_nested_switch_statement
# A switch statement nested inside an outer switch clause action block is properly located
$code = @"
switch (`$category) {
    'hardware' {
        switch (`$subCategory) {
            'cpu' { 1 }
            'ram' { 2 }
        }
    }
}
"@

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$switches = @($ast.FindAll({ $args[0] -is [System.Management.Automation.Language.SwitchStatementAst] }, $true))

if ($switches.Count -ne 2) {
    Write-Host "FAIL: expected 2 SwitchStatementAst instances, got $($switches.Count)"
    exit 1
}

$outer = $switches[0]
$inner = $switches[1]

if ($outer.Condition.Extent.Text -ne "`$category" -or $inner.Condition.Extent.Text -ne "`$subCategory") {
    Write-Host "FAIL: outer/inner conditions mismatch"
    exit 1
}

Write-Host "PASS"
exit 0
