# vybe-test: powershell/ast_string_constant_expression_ast/ast_string_constant_parent_hierarchy
# StringConstantExpressionAst.Parent properly links to the enclosing AST element
$code = '$lookup = @{ key = ''value'' }'

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$keyNode = $ast.Find({ $args[0] -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $args[0].Value -eq "key" }, $true)

if ($keyNode -eq $null) {
    Write-Host "FAIL: key node not found"
    exit 1
}

$parent = $keyNode.Parent

if ($parent -isnot [System.Management.Automation.Language.HashtableAst]) {
    Write-Host "FAIL: expected parent of hashtable key to be HashtableAst, got '$($parent.GetType().Name)'"
    exit 1
}

Write-Host "PASS"
exit 0
