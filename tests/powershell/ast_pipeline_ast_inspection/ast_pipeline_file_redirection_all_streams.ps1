# vybe-test: powershell/ast_pipeline_ast_inspection/ast_pipeline_file_redirection_all_streams
# *> all.log parses as FileRedirectionAst with FromStream = All
$code = "Build-Solution *> all_streams.log"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$cmdAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true)
$redir = $cmdAst.Redirections[0]

if ($redir.FromStream -ne [System.Management.Automation.Language.RedirectionStream]::All) {
    Write-Host "FAIL: FromStream mismatch, expected All, got '$($redir.FromStream)'"
    exit 1
}

Write-Host "PASS"
exit 0
