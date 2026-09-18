# vybe-test: powershell/ast_symbol_resolution/ast_symbol_variable_path_script_scope
# Variable with $script: prefix sets IsScript to true on VariablePath
$code = "`$script:InternalCache = @{}"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$varAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.VariableExpressionAst] }, $true)
$vp = $varAst.VariablePath

if (-not $vp.IsScript) {
    Write-Host "FAIL: IsScript was false for script variable"
    exit 1
}

if ($vp.UserPath -ne "script:InternalCache") {
    Write-Host "FAIL: UserPath mismatch: '$($vp.UserPath)'"
    exit 1
}

Write-Host "PASS"
exit 0
