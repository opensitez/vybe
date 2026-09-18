# vybe-test: powershell/format_list_engine/list_dictionary_generic_rendering
$dict = [System.Collections.Generic.Dictionary[string, string]]::new()
$dict["Endpoint"] = "api.example.com"

# Generic Dictionary[string, string] formatted in list view
$output = $dict | Format-List | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

if (-not ($output -match "Key\s*:\s*Endpoint" -and $output -match "Value\s*:\s*api\.example\.com")) {
    Write-Host "FAIL: generic dictionary failed to format in list view: $output"
    exit 1
}

Write-Host "PASS"
exit 0
