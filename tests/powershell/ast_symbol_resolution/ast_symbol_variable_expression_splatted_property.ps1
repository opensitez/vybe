# vybe-test: powershell/ast_symbol_resolution/ast_symbol_variable_expression_splatted_property
# VariableExpressionAst.Splatted is true for @splat and false for $normal
$code = @"
`$normalArgs = @{ Path = 'test' }
Get-ChildItem @normalArgs
"@

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$vars = @($ast.FindAll({ $args[0] -is [System.Management.Automation.Language.VariableExpressionAst] }, $true))

$regular = $vars | Where-Object { -not $_.Splatted }
$splatted = $vars | Where-Object { $_.Splatted }

if ($regular -eq $null) {
    Write-Host "FAIL: regular un-splatted variable expression missing"
    exit 1
}

if ($splatted -eq $null) {
    Write-Host "FAIL: splatted variable expression missing"
    exit 1
}

if ($splatted.VariablePath.UserPath -ne "normalArgs") {
    Write-Host "FAIL: splatted variable UserPath mismatch: '$($splatted.VariablePath.UserPath)'"
    exit 1
}

Write-Host "PASS"
exit 0
