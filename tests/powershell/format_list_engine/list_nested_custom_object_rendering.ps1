# vybe-test: powershell/format_list_engine/list_nested_custom_object_rendering
$nested = [pscustomobject]@{
    OuterName = "ParentContainer"
    InnerObj  = [pscustomobject]@{ SubField = "SubValue" }
}

# Nested PSCustomObjects are formatted compactly as @{SubField=SubValue} in list view
$output = $nested | Format-List | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

if (-not ($output -match "OuterName\s*:\s*ParentContainer" -and $output -match "InnerObj\s*:\s*@\{SubField=SubValue\}")) {
    Write-Host "FAIL: nested custom object failed to render in list view: $output"
    exit 1
}

Write-Host "PASS"
exit 0
