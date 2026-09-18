# vybe-test: powershell/supports_paging_parameters/supports_paging_explicit_false_omits_paging_parameters
# Setting [CmdletBinding(SupportsPaging = $false)] does not add First, Skip, or IncludeTotalCount
function TestExplicitFalsePaging {
    [CmdletBinding(SupportsPaging = $false)]
    param()
}

$params = (Get-Command TestExplicitFalsePaging).Parameters

if ($params.ContainsKey("First") -or $params.ContainsKey("Skip") -or $params.ContainsKey("IncludeTotalCount")) {
    Write-Host "FAIL: paging parameters were unexpectedly added under SupportsPaging = `$false"
    exit 1
}

Write-Host "PASS"
exit 0
