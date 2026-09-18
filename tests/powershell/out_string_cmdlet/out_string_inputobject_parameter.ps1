# vybe-test: powershell/out_string_cmdlet/out_string_inputobject_parameter
# Out-String accepts items directly through the -InputObject parameter
$output = Out-String -InputObject "standalone_input_object"

if ($output -notmatch "standalone_input_object") {
    Write-Host "FAIL: expected 'standalone_input_object' in output, got: '$output'"
    exit 1
}

Write-Host "PASS"
exit 0
