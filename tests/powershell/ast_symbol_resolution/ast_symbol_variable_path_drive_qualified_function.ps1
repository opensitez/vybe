# vybe-test: powershell/ast_symbol_resolution/ast_symbol_variable_path_drive_qualified_function
# Variable with $function: prefix sets DriveName to 'function' and IsDriveQualified to true
$code = "`$function:custom_prompt = { 'PS> ' }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$varAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.VariableExpressionAst] }, $true)
$vp = $varAst.VariablePath

if ($vp.DriveName -ne "function") {
    Write-Host "FAIL: DriveName was not 'function': '$($vp.DriveName)'"
    exit 1
}

if (-not $vp.IsDriveQualified) {
    Write-Host "FAIL: IsDriveQualified was false for function drive"
    exit 1
}

Write-Host "PASS"
exit 0
