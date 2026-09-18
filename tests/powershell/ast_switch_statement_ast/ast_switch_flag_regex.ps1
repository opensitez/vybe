# vybe-test: powershell/ast_switch_statement_ast/ast_switch_flag_regex
# switch -Regex ($val) sets SwitchStatementAst.Flags to include SwitchFlags.Regex
$code = "switch -Regex (`$line) { '^Error' { 1 } }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$switchAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.SwitchStatementAst] }, $true)

$hasRegex = ($switchAst.Flags -band [System.Management.Automation.Language.SwitchFlags]::Regex) -ne 0

if (-not $hasRegex) {
    Write-Host "FAIL: Regex flag was not set, Flags: $($switchAst.Flags)"
    exit 1
}

Write-Host "PASS"
exit 0
