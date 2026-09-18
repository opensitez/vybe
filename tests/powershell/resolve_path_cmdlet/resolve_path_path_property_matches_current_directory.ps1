# vybe-test: powershell/resolve_path_cmdlet/resolve_path_path_property_matches_current_directory
# The .Path property of the returned PathInfo matches the absolute working directory
$info = Resolve-Path "."

if ($info.Path -ne $PWD.Path) {
    Write-Host "FAIL: expected '$($PWD.Path)', got: '$($info.Path)'"
    exit 1
}

Write-Host "PASS"
exit 0
