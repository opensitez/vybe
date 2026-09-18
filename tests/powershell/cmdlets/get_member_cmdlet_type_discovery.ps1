# vybe-test: powershell/cmdlets/get_member_cmdlet_type_discovery
# Get-Member reflects the available methods and properties on pipeline input objects
$stringObj = "PowershellPlatform"

$members = $stringObj | Get-Member -MemberType Method -Name "Substring"

if ($members -eq $null) {
    Write-Host "FAIL: Get-Member did not discover Substring method"
    exit 1
}

if ($members.Name -ne "Substring") {
    Write-Host "FAIL: expected member name 'Substring', got '$($members.Name)'"
    exit 1
}

$prop = $stringObj | Get-Member -MemberType Property -Name "Length"
if ($prop -eq $null -or $prop.Name -ne "Length") {
    Write-Host "FAIL: Get-Member did not discover Length property"
    exit 1
}

Write-Host "PASS"
exit 0
