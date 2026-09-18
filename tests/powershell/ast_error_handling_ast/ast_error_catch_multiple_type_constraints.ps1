# vybe-test: powershell/ast_error_handling_ast/ast_error_catch_multiple_type_constraints
# CatchClauseAst handles multiple exception types separated by commas
$code = "try { 1 } catch [System.IO.FileNotFoundException], [System.IO.DirectoryNotFoundException] { 2 }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$catchAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.CatchClauseAst] }, $true)

if ($catchAst.CatchTypes.Count -ne 2) {
    Write-Host "FAIL: expected 2 catch types, got $($catchAst.CatchTypes.Count)"
    exit 1
}

$first = $catchAst.CatchTypes[0].TypeName.FullName
$second = $catchAst.CatchTypes[1].TypeName.FullName

if ($first -ne "System.IO.FileNotFoundException" -or $second -ne "System.IO.DirectoryNotFoundException") {
    Write-Host "FAIL: catch types mismatch: '$first', '$second'"
    exit 1
}

Write-Host "PASS"
exit 0
