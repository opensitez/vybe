# vybe-test: powershell/ast_switch_statement_ast/ast_switch_flag_combined_regex_casesensitive
# switch -Regex -CaseSensitive combines both flags into SwitchStatementAst.Flags bitmask
$code = "switch -Regex -CaseSensitive (`$text) { '^[A-Z]+$' { 'ALL_CAPS' } }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$switchAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.SwitchStatementAst] }, $true)

$hasRegex = ($switchAst.Flags -band [System.Management.Automation.Language.SwitchFlags]::Regex) -ne 0
$hasCase = ($switchAst.Flags -band [System.Management.Automation.Language.SwitchFlags]::CaseSensitive) -ne 0

if (-not $hasRegex -or -not $hasCase) {
    Write-Host "FAIL: combined flags mismatch, Flags: $($switchAst.Flags)"
    exit 1
}

Write-Host "PASS"
exit 0
