# vybe-test: powershell/ast_type_expression_ast/ast_type_expression_static_property_access
# In '[System.DateTime]::Now', the member expression target is a TypeExpressionAst
$code = "`$currentTime = [System.DateTime]::Now"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$memberAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.MemberExpressionAst] }, $true)

$targetExpr = $memberAst.Expression

if ($targetExpr -isnot [System.Management.Automation.Language.TypeExpressionAst]) {
    Write-Host "FAIL: target expression was not TypeExpressionAst"
    exit 1
}

if ($targetExpr.TypeName.FullName -ne "System.DateTime") {
    Write-Host "FAIL: target TypeName mismatch: '$($targetExpr.TypeName.FullName)'"
    exit 1
}

if ($memberAst.Member.Value -ne "Now") {
    Write-Host "FAIL: member property mismatch: '$($memberAst.Member.Value)'"
    exit 1
}

Write-Host "PASS"
exit 0
