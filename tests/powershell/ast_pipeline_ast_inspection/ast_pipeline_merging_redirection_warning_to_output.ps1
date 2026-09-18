# vybe-test: powershell/ast_pipeline_ast_inspection/ast_pipeline_merging_redirection_warning_to_output
# 3>&1 parses as MergingRedirectionAst with FromStream = Warning and ToStream = Output
$code = "Invoke-Task 3>&1"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$cmdAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true)
$redir = $cmdAst.Redirections[0]

if ($redir -isnot [System.Management.Automation.Language.MergingRedirectionAst]) {
    Write-Host "FAIL: expected MergingRedirectionAst"
    exit 1
}

if ($redir.FromStream -ne [System.Management.Automation.Language.RedirectionStream]::Warning) {
    Write-Host "FAIL: FromStream mismatch, expected Warning, got '$($redir.FromStream)'"
    exit 1
}

if ($redir.ToStream -ne [System.Management.Automation.Language.RedirectionStream]::Output) {
    Write-Host "FAIL: ToStream mismatch, expected Output, got '$($redir.ToStream)'"
    exit 1
}

Write-Host "PASS"
exit 0
