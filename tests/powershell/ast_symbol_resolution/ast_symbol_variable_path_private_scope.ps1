# vybe-test: powershell/ast_symbol_resolution/ast_symbol_variable_path_private_scope
# Variable with $private: prefix sets IsPrivate to true on VariablePath
$code = "`$private:isolatedToken = 'SECRET'"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$varAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.VariableExpressionAst] }, $true)
$vp = $varAst.VariablePath

if (-not $vp.IsPrivate) {
    Write-Host "FAIL: IsPrivate was false for private variable"
    exit 1
}

if ($vp.UserPath -ne "private:isolatedToken") {
    Write-Host "FAIL: UserPath mismatch: '$($vp.UserPath)'"
    exit 1
}

Write-Host "PASS"
exit 0
