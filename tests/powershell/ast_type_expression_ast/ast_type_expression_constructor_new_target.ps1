# vybe-test: powershell/ast_type_expression_ast/ast_type_expression_constructor_new_target
# In '[System.Text.StringBuilder]::new()', the invocation target is a TypeExpressionAst
$code = "`$sb = [System.Text.StringBuilder]::new()"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$invokeAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.InvokeMemberExpressionAst] }, $true)

$targetExpr = $invokeAst.Expression

if ($targetExpr -isnot [System.Management.Automation.Language.TypeExpressionAst]) {
    Write-Host "FAIL: constructor target was not TypeExpressionAst"
    exit 1
}

if ($targetExpr.TypeName.FullName -ne "System.Text.StringBuilder") {
    Write-Host "FAIL: constructor target TypeName mismatch: '$($targetExpr.TypeName.FullName)'"
    exit 1
}

if ($invokeAst.Member.Value -ne "new") {
    Write-Host "FAIL: member name mismatch: '$($invokeAst.Member.Value)'"
    exit 1
}

Write-Host "PASS"
exit 0
