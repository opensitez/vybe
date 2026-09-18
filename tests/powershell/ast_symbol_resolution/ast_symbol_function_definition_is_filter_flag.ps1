# vybe-test: powershell/ast_symbol_resolution/ast_symbol_function_definition_is_filter_flag
# FunctionDefinitionAst.IsFilter is true for functions declared using the 'filter' keyword
$code = @"
function NormalFunc { `$true }
filter StreamingFilter { `$_ * 10 }
"@

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$funcs = @($ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true))

$normal = $funcs | Where-Object { $_.Name -eq "NormalFunc" }
$filter = $funcs | Where-Object { $_.Name -eq "StreamingFilter" }

if ($normal.IsFilter) {
    Write-Host "FAIL: IsFilter was unexpectedly true for normal function"
    exit 1
}

if (-not $filter.IsFilter) {
    Write-Host "FAIL: IsFilter was false for filter function"
    exit 1
}

Write-Host "PASS"
exit 0
