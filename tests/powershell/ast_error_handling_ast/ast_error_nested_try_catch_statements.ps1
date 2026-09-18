# vybe-test: powershell/ast_error_handling_ast/ast_error_nested_try_catch_statements
# A TryStatementAst nested within another TryStatementAst is properly resolved via AST search
$code = @"
try {
    try {
        1
    } catch {
        2
    }
} catch {
    3
}
"@

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$tryList = @($ast.FindAll({ $args[0] -is [System.Management.Automation.Language.TryStatementAst] }, $true))

if ($tryList.Count -ne 2) {
    Write-Host "FAIL: expected 2 TryStatementAst instances, got $($tryList.Count)"
    exit 1
}

$outerTry = $tryList[0]
$innerTry = $tryList[1]

# Inner try must be a descendant of outer try
$isDescendant = $false
$curr = $innerTry.Parent
while ($curr -ne $null) {
    if ($curr -eq $outerTry) {
        $isDescendant = $true
        break
    }
    $curr = $curr.Parent
}

if (-not $isDescendant) {
    Write-Host "FAIL: inner try was not a descendant of outer try"
    exit 1
}

Write-Host "PASS"
exit 0
