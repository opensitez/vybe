# vybe-test: powershell/ast_type_expression_ast/ast_type_expression_static_method_target
# In static method call '[Math]::Sqrt(16)', the target expression is a TypeExpressionAst
$code = "`$res = [Math]::Sqrt(16)"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$invokeMemberAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.InvokeMemberExpressionAst] }, $true)

$targetExpr = $invokeMemberAst.Expression

if ($targetExpr -isnot [System.Management.Automation.Language.TypeExpressionAst]) {
    Write-Host "FAIL: target expression was not TypeExpressionAst, got '$($targetExpr.GetType().Name)'"
    exit 1
}

if ($targetExpr.TypeName.FullName -ne "Math") {
    Write-Host "FAIL: target TypeName mismatch: '$($targetExpr.TypeName.FullName)'"
    exit 1
}

Write-Host "PASS"
exit 0
