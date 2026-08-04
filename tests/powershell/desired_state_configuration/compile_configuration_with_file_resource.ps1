# vybe-test: powershell/desired_state_configuration/compile_configuration_with_file_resource
configuration FileConfig {
    Node 'localhost' {
        File ExampleFile {
            DestinationPath = "$PWD/dsc-file.txt"
            Contents = 'hello'
        }
    }
}
FileConfig -OutputPath "$PWD/dsc-file"
if (-not (Test-Path "$PWD/dsc-file/localhost.mof")) {
    Write-Host "FAIL: expected file MOF"
    exit 1
}
Remove-Item -Recurse -Force "$PWD/dsc-file"
Write-Host 'PASS'
exit 0
