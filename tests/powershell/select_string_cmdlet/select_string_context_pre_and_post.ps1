# vybe-test: powershell/select_string_cmdlet/select_string_context_pre_and_post
# -Context captures lines before (PreContext) and after (PostContext) the match
$lines = @("line 1 before", "line 2 target match", "line 3 after")
$match = $lines | Select-String -Pattern "target" -Context 1

if ($null -eq $match.Context) {
    Write-Host "FAIL: Context property missing on MatchInfo"
    exit 1
}

if ($match.Context.PreContext[0] -ne "line 1 before") {
    Write-Host "FAIL: PreContext mismatch, got: '$($match.Context.PreContext[0])'"
    exit 1
}

if ($match.Context.PostContext[0] -ne "line 3 after") {
    Write-Host "FAIL: PostContext mismatch, got: '$($match.Context.PostContext[0])'"
    exit 1
}

Write-Host "PASS"
exit 0
