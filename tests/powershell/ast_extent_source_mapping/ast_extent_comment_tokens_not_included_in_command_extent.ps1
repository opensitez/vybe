# vybe-test: powershell/ast_extent_source_mapping/ast_extent_comment_tokens_not_included_in_command_extent
# Trailing comments on the same line are excluded from a statement's AST extent
$code = "`$value = 999  # this is a trailing explanation comment"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$assignAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true)
$assignText = $assignAst.Extent.Text

if ($assignText -ne "`$value = 999") {
    Write-Host "FAIL: assignment extent unexpectedly included trailing comment: '$assignText'"
    exit 1
}

# The comment token exists independently in the token stream
$commentToken = $tokens | Where-Object { $_.Kind -eq [System.Management.Automation.Language.TokenKind]::Comment }
if ($commentToken -eq $null -or $commentToken.Text -ne "# this is a trailing explanation comment") {
    Write-Host "FAIL: comment token missing or text mismatch"
    exit 1
}

Write-Host "PASS"
exit 0
