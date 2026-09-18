# vybe-test: powershell/ast_pipeline_ast_inspection/ast_pipeline_merging_redirection_error_to_output
# 2>&1 parses as MergingRedirectionAst with FromStream = Error and ToStream = Output
$code = "Get-Process -Name 'missing' 2>&1"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$cmdAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true)

if ($cmdAst.Redirections.Count -ne 1) {
    Write-Host "FAIL: expected 1 redirection, got $($cmdAst.Redirections.Count)"
    exit 1
}

$redir = $cmdAst.Redirections[0]

if ($redir -isnot [System.Management.Automation.Language.MergingRedirectionAst]) {
    Write-Host "FAIL: expected MergingRedirectionAst, got '$($redir.GetType().Name)'"
    exit 1
}

if ($redir.FromStream -ne [System.Management.Automation.Language.RedirectionStream]::Error) {
    Write-Host "FAIL: FromStream mismatch, expected Error, got '$($redir.FromStream)'"
    exit 1
}

if ($redir.ToStream -ne [System.Management.Automation.Language.RedirectionStream]::Output) {
    Write-Host "FAIL: ToStream mismatch, expected Output, got '$($redir.ToStream)'"
    exit 1
}

Write-Host "PASS"
exit 0
