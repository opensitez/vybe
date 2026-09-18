# vybe-test: powershell/ast_switch_statement_ast/ast_switch_flag_file
# switch -File $path sets SwitchStatementAst.Flags to include SwitchFlags.File
$code = "switch -File 'input.log' { 'ERROR' { 1 } }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$switchAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.SwitchStatementAst] }, $true)

$hasFile = ($switchAst.Flags -band [System.Management.Automation.Language.SwitchFlags]::File) -ne 0

if (-not $hasFile) {
    Write-Host "FAIL: File flag was not set, Flags: $($switchAst.Flags)"
    exit 1
}

Write-Host "PASS"
exit 0
