# vybe-test: powershell/parameters_alias_attribute/multiple_aliases_on_single_parameter
function Set-DestPath {
    param(
        [Alias("Dest", "TargetPath", "OutPath")]
        [string]$Destination
    )
    return "Dest:$Destination"
}
$r1 = Set-DestPath -Dest "/tmp/a"
$r2 = Set-DestPath -TargetPath "/tmp/b"
$r3 = Set-DestPath -OutPath "/tmp/c"
if ($r1 -ne "Dest:/tmp/a" -or $r2 -ne "Dest:/tmp/b" -or $r3 -ne "Dest:/tmp/c") {
    Write-Host "FAIL: Multiple parameter aliases failed"
    exit 1
}
Write-Host "PASS"
exit 0
