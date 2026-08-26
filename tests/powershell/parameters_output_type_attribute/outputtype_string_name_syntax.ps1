# vybe-test: powershell/parameters_output_type_attribute/outputtype_string_name_syntax
function Get-NamedTypeOutput {
    [OutputType("System.Guid")]
    param()
    return [guid]::NewGuid()
}
$cmd = Get-Command Get-NamedTypeOutput
$types = @($cmd.OutputType | ForEach-Object { $_.Name })
if ($types -notcontains "System.Guid") {
    Write-Host "FAIL: OutputType string name syntax failed"
    exit 1
}
Write-Host "PASS"
exit 0
