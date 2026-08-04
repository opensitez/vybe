# vybe-test: powershell/comment_based_help/help_with_parameters
<##
.PARAMETER Name
Name parameter.
#>
function Test-Help { param($Name) }
Write-Host 'PASS'
exit 0
