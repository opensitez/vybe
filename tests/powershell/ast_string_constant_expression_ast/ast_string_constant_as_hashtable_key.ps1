# vybe-test: powershell/ast_string_constant_expression_ast/ast_string_constant_as_hashtable_key
# Hashtable string keys parse as StringConstantExpressionAst
$code = "`$table = @{ 'Content-Type' = 'application/json' }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$hashAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.HashtableAst] }, $true)
$pair = $hashAst.KeyValuePairs[0]

$keyExpr = $pair.Item1

if ($keyExpr -isnot [System.Management.Automation.Language.StringConstantExpressionAst]) {
    Write-Host "FAIL: key was not StringConstantExpressionAst"
    exit 1
}

if ($keyExpr.Value -ne "Content-Type") {
    Write-Host "FAIL: key value mismatch: '$($keyExpr.Value)'"
    exit 1
}

Write-Host "PASS"
exit 0
