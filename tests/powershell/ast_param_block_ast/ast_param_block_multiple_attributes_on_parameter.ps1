# vybe-test: powershell/ast_param_block_ast/ast_param_block_multiple_attributes_on_parameter
# ParameterAst.Attributes contains all parameter decorators and type constraints
$code = "param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]`$Database)"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$paramAst = $ast.ParamBlock.Parameters[0]

# Expected: Parameter attribute, ValidateNotNullOrEmpty attribute, and string type constraint
if ($paramAst.Attributes.Count -ne 3) {
    Write-Host "FAIL: expected 3 attributes, got $($paramAst.Attributes.Count)"
    exit 1
}

$names = @($paramAst.Attributes | ForEach-Object { $_.TypeName.FullName })

if (-not $names.Contains("Parameter") -or -not $names.Contains("ValidateNotNullOrEmpty") -or -not $names.Contains("string")) {
    Write-Host "FAIL: attribute names mismatch: $($names -join ', ')"
    exit 1
}

Write-Host "PASS"
exit 0
