# vybe-test: powershell/ast_symbol_resolution/ast_symbol_ast_parent_hierarchy_traversal
# Traversing the Ast.Parent hierarchy from a leaf node successfully ascends to the root ScriptBlockAst
$code = "function Outer { `$leafVariable = 42 }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$leafNode = $ast.Find({ $args[0] -is [System.Management.Automation.Language.VariableExpressionAst] }, $true)

$current = $leafNode
$path = @()
while ($current -ne $null) {
    $path += $current.GetType().Name
    $current = $current.Parent
}

# The top root node must be ScriptBlockAst
$rootType = $path[-1]
if ($rootType -ne "ScriptBlockAst") {
    Write-Host "FAIL: top root of parent hierarchy was not ScriptBlockAst, got '$rootType'"
    exit 1
}

# Path must contain FunctionDefinitionAst
if (-not $path.Contains("FunctionDefinitionAst")) {
    Write-Host "FAIL: FunctionDefinitionAst not encountered in parent hierarchy"
    exit 1
}

Write-Host "PASS"
exit 0
