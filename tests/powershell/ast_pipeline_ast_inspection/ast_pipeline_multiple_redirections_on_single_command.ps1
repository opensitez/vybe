# vybe-test: powershell/ast_pipeline_ast_inspection/ast_pipeline_multiple_redirections_on_single_command
# A command with multiple redirections (2> err.txt > out.txt) parses both in CommandAst.Redirections
$code = "Start-Build 2> 'errors.txt' > 'output.txt'"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$cmdAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true)

if ($cmdAst.Redirections.Count -ne 2) {
    Write-Host "FAIL: expected 2 redirections, got $($cmdAst.Redirections.Count)"
    exit 1
}

$hasErrorRedir = ($cmdAst.Redirections | Where-Object { $_.FromStream -eq [System.Management.Automation.Language.RedirectionStream]::Error -and $_.Location.Value -eq "errors.txt" }) -ne $null
$hasOutputRedir = ($cmdAst.Redirections | Where-Object { $_.FromStream -eq [System.Management.Automation.Language.RedirectionStream]::Output -and $_.Location.Value -eq "output.txt" }) -ne $null

if (-not $hasErrorRedir) {
    Write-Host "FAIL: Error stream redirection not found in Redirections"
    exit 1
}

if (-not $hasOutputRedir) {
    Write-Host "FAIL: Output stream redirection not found in Redirections"
    exit 1
}

Write-Host "PASS"
exit 0
