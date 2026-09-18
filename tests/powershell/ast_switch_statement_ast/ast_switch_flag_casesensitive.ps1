# vybe-test: powershell/ast_switch_statement_ast/ast_switch_flag_casesensitive
# switch -CaseSensitive ($val) sets SwitchStatementAst.Flags to include SwitchFlags.CaseSensitive
$code = "switch -CaseSensitive (`$code) { 'DEBUG' { 1 } }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$switchAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.SwitchStatementAst] }, $true)

$hasCase = ($switchAst.Flags -band [System.Management.Automation.Language.SwitchFlags]::CaseSensitive) -ne 0

if (-not $hasCase) {
    Write-Host "FAIL: CaseSensitive flag was not set, Flags: $($switchAst.Flags)"
    exit 1
}

Write-Host "PASS"
exit 0
