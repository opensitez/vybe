# vybe-test: powershell/parameters_alias_attribute/alias_reflection_inspection_in_command_metadata
function Target-FuncWithAlias {
    param(
        [Alias("Alt1", "Alt2")]
        [string]$Original
    )
}
$cmd = Get-Command Target-FuncWithAlias
$aliases = @($cmd.Parameters["Original"].Aliases)
if ($aliases.Length -ne 2 -or -not ($aliases -contains "Alt1") -or -not ($aliases -contains "Alt2")) {
    Write-Host "FAIL: Parameter alias reflection inspection failed"
    exit 1
}
Write-Host "PASS"
exit 0
