# vybe-test: powershell/parameters_output_type_attribute/outputtype_custom_class_type
class UserProfileData {
    [string]$Username
}
function Get-UserProfile {
    [OutputType([UserProfileData])]
    param()
    $u = [UserProfileData]::new()
    $u.Username = "alice"
    return $u
}
$cmd = Get-Command Get-UserProfile
$types = @($cmd.OutputType | ForEach-Object { $_.Type.Name })
if ($types -notcontains "UserProfileData") {
    Write-Host "FAIL: OutputType custom class metadata failed"
    exit 1
}
Write-Host "PASS"
exit 0
