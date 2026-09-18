# vybe-test: powershell/supports_paging_parameters/supports_paging_adds_paging_parameters_to_command
# [CmdletBinding(SupportsPaging = $true)] automatically adds First, Skip, and IncludeTotalCount parameters
function TestPagingExposed {
    [CmdletBinding(SupportsPaging = $true)]
    param()
}

$cmd = Get-Command TestPagingExposed
$params = $cmd.Parameters

if (-not $params.ContainsKey("First")) {
    Write-Host "FAIL: 'First' parameter not found"
    exit 1
}

if (-not $params.ContainsKey("Skip")) {
    Write-Host "FAIL: 'Skip' parameter not found"
    exit 1
}

if (-not $params.ContainsKey("IncludeTotalCount")) {
    Write-Host "FAIL: 'IncludeTotalCount' parameter not found"
    exit 1
}

Write-Host "PASS"
exit 0
