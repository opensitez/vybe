# vybe-test: powershell/parameters_validate_length/validatelength_with_splatted_hashtable
function Set-Title {
    param([ValidateLength(3, 20)][string]$Title)
    return "Title:$Title"
}
$params = @{ Title = "Chapter 1" }
$res = Set-Title @params
if ($res -ne "Title:Chapter 1") {
    Write-Host "FAIL: ValidateLength splatting failed"
    exit 1
}
Write-Host "PASS"
exit 0
