# vybe-test: powershell/help_comments/comments_with_parameter
<#
.PARAMETER Name
Name parameter.
#>
function Test-HelpCommentParameter { param($Name) }
Write-Host 'PASS'
exit 0
