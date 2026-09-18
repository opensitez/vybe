# vybe-test: powershell/ast_loop_statements/ast_loop_for_multiple_initializers
# For loop supports multiple comma-separated statements in Initializer and Iterator
$code = "for (`$i = 0, `$j = 10; `$i -lt `$j; `$i++, `$j--) { `$i + `$j }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$forAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.ForStatementAst] }, $true)

$initText = $forAst.Initializer.Extent.Text
$iterText = $forAst.Iterator.Extent.Text

if ($initText -ne "`$i = 0, `$j = 10") {
    Write-Host "FAIL: multiple initializers text mismatch: '$initText'"
    exit 1
}

if ($iterText -ne "`$i++, `$j--") {
    Write-Host "FAIL: multiple iterators text mismatch: '$iterText'"
    exit 1
}

Write-Host "PASS"
exit 0
