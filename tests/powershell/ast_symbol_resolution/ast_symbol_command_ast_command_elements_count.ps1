# vybe-test: powershell/ast_symbol_resolution/ast_symbol_command_ast_command_elements_count
# CommandAst.CommandElements contains the command name and all parameter/argument AST expressions
$code = "Select-Object -Property Name, Id -First 10"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$cmdAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true)

# Elements: Select-Object, -Property, (Name, Id), -First, 10
if ($cmdAst.CommandElements.Count -ne 5) {
    Write-Host "FAIL: expected 5 command elements, got $($cmdAst.CommandElements.Count)"
    exit 1
}

$firstElem = $cmdAst.CommandElements[0].Extent.Text
if ($firstElem -ne "Select-Object") {
    Write-Host "FAIL: first command element was '$firstElem'"
    exit 1
}

Write-Host "PASS"
exit 0
