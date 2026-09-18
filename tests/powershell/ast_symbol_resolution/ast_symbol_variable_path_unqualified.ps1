# vybe-test: powershell/ast_symbol_resolution/ast_symbol_variable_path_unqualified
# Simple variable without scope prefix evaluates IsUnqualified as true and DriveName as null
$code = "`$standardVariable = 'payload'"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$varAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.VariableExpressionAst] }, $true)
$vp = $varAst.VariablePath

if (-not $vp.IsUnqualified) {
    Write-Host "FAIL: IsUnqualified was false for simple variable"
    exit 1
}

if ($vp.DriveName -ne $null) {
    Write-Host "FAIL: DriveName was not null: '$($vp.DriveName)'"
    exit 1
}

if ($vp.IsGlobal -or $vp.IsScript -or $vp.IsLocal -or $vp.IsPrivate) {
    Write-Host "FAIL: scope flags were unexpectedly set on unqualified variable"
    exit 1
}

Write-Host "PASS"
exit 0
