# vybe-test: powershell/ast_param_block_ast/ast_param_block_nested_function_param_block
# Finding the ParamBlockAst within a nested function definition resolves its isolated parameters
$code = @"
function Invoke-NestedProcess {
    param([int]`$OuterLimit = 100)
    
    function HelperFunction {
        param([string]`$HelperMessage)
        Write-Host `$HelperMessage
    }
}
"@

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$funcs = @($ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true))
$helperFunc = $funcs | Where-Object { $_.Name -eq "HelperFunction" }

$helperParamBlock = $helperFunc.Body.ParamBlock

if ($helperParamBlock -eq $null) {
    Write-Host "FAIL: helper function ParamBlock was null"
    exit 1
}

$helperParamName = $helperParamBlock.Parameters[0].Name.VariablePath.UserPath

if ($helperParamName -ne "HelperMessage") {
    Write-Host "FAIL: helper parameter name mismatch: '$helperParamName'"
    exit 1
}

Write-Host "PASS"
exit 0
