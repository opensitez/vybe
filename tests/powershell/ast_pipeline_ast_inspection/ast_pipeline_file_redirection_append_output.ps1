# vybe-test: powershell/ast_pipeline_ast_inspection/ast_pipeline_file_redirection_append_output
# >> file.txt parses as FileRedirectionAst with Append = true and FromStream = Output
$code = "Get-Date >> execution.log"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$cmdAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true)
$redir = $cmdAst.Redirections[0]

if ($redir -isnot [System.Management.Automation.Language.FileRedirectionAst]) {
    Write-Host "FAIL: expected FileRedirectionAst"
    exit 1
}

if (-not $redir.Append) {
    Write-Host "FAIL: Append was false for >> append redirection"
    exit 1
}

if ($redir.FromStream -ne [System.Management.Automation.Language.RedirectionStream]::Output) {
    Write-Host "FAIL: FromStream mismatch, expected Output, got '$($redir.FromStream)'"
    exit 1
}

Write-Host "PASS"
exit 0
