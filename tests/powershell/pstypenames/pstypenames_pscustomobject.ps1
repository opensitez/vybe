# vybe-test: powershell/pstypenames/pstypenames_pscustomobject
$obj = [pscustomobject]@{ Item = "Test" }
$types = @($obj.psobject.TypeNames)
if ($types -notcontains "System.Management.Automation.PSCustomObject") {
    Write-Host "FAIL: PSCustomObject default TypeNames missing PSCustomObject type"
    exit 1
}
Write-Host "PASS"
exit 0
