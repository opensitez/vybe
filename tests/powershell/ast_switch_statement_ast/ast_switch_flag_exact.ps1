# vybe-test: powershell/ast_switch_statement_ast/ast_switch_flag_exact
# switch -Exact ($val) sets SwitchStatementAst.Flags to SwitchFlags.Exact or None (without Regex/Wildcard)
$code = "switch -Exact (`$code) { 10 { 'ten' } }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$switchAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.SwitchStatementAst] }, $true)

$hasRegex = ($switchAst.Flags -band [System.Management.Automation.Language.SwitchFlags]::Regex) -ne 0
$hasWildcard = ($switchAst.Flags -band [System.Management.Automation.Language.SwitchFlags]::Wildcard) -ne 0

if ($hasRegex -or $hasWildcard) {
    Write-Host "FAIL: exact switch had Regex or Wildcard flags set"
    exit 1
}

Write-Host "PASS"
exit 0
