# vybe-test: powershell/variables/new_variable_description_metadata
# New-Variable stores description metadata on the created PSVariable object
$varName = "metadataVar_$PID"

try {
    New-Variable -Name $varName -Value 999 -Description "Custom audit metadata string" -Force
    $varObj = Get-Variable -Name $varName

    if ($varObj.Description -ne "Custom audit metadata string") {
        Write-Host "FAIL: expected description 'Custom audit metadata string', got '$($varObj.Description)'"
        exit 1
    }

    if ($varObj.Value -ne 999) {
        Write-Host "FAIL: expected value 999, got $($varObj.Value)"
        exit 1
    }
} finally {
    Remove-Variable -Name $varName -Force -ErrorAction SilentlyContinue
}

Write-Host "PASS"
exit 0
