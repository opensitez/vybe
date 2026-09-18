# vybe-test: powershell/ast_symbol_resolution/ast_symbol_variable_path_drive_qualified_env
# Variable with $env: prefix sets DriveName to 'env' and IsDriveQualified to true
$code = "`$env:SYSTEM_PATH = '/usr/bin'"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$varAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.VariableExpressionAst] }, $true)
$vp = $varAst.VariablePath

if ($vp.DriveName -ne "env") {
    Write-Host "FAIL: DriveName was not 'env': '$($vp.DriveName)'"
    exit 1
}

if (-not $vp.IsDriveQualified) {
    Write-Host "FAIL: IsDriveQualified was false for env variable"
    exit 1
}

Write-Host "PASS"
exit 0
