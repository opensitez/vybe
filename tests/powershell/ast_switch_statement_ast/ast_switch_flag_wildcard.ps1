# vybe-test: powershell/ast_switch_statement_ast/ast_switch_flag_wildcard
# switch -Wildcard ($val) sets SwitchStatementAst.Flags to include SwitchFlags.Wildcard
$code = "switch -Wildcard (`$path) { '*.ps1' { 1 } }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$switchAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.SwitchStatementAst] }, $true)

$hasWildcard = ($switchAst.Flags -band [System.Management.Automation.Language.SwitchFlags]::Wildcard) -ne 0

if (-not $hasWildcard) {
    Write-Host "FAIL: Wildcard flag was not set, Flags: $($switchAst.Flags)"
    exit 1
}

Write-Host "PASS"
exit 0
