# vybe-test: powershell/ast_switch_statement_ast/ast_switch_multiple_pattern_types_in_clauses
# A switch statement can mix string literals, numeric constants, and scriptblocks across its clauses
$code = @"
switch (`$val) {
    'literal'             { 'matched string' }
    100                   { 'matched number' }
    { `$_ -is [string] }   { 'matched type' }
}
"@

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$switchAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.SwitchStatementAst] }, $true)

$type1 = $switchAst.Clauses[0].Item1.GetType().Name
$type2 = $switchAst.Clauses[1].Item1.GetType().Name
$type3 = $switchAst.Clauses[2].Item1.GetType().Name

if ($type1 -ne "StringConstantExpressionAst") {
    Write-Host "FAIL: clause 0 type mismatch: '$type1'"
    exit 1
}

if ($type2 -ne "ConstantExpressionAst") {
    Write-Host "FAIL: clause 1 type mismatch: '$type2'"
    exit 1
}

if ($type3 -ne "ScriptBlockExpressionAst") {
    Write-Host "FAIL: clause 2 type mismatch: '$type3'"
    exit 1
}

Write-Host "PASS"
exit 0
