# vybe-test: powershell/format_list_engine/list_heterogeneous_objects
$items = @(
    [pscustomobject]@{ TypeOne = "Val1" },
    [pscustomobject]@{ TypeTwo = "Val2" }
)

# In Format-List, heterogeneous objects display with their own property sets
$output = $items | Format-List -Property TypeOne, TypeTwo | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

if (-not ($output -match "TypeOne\s*:\s*Val1" -and $output -match "TypeTwo\s*:\s*Val2")) {
    Write-Host "FAIL: heterogeneous objects failed to format in list view: $output"
    exit 1
}

Write-Host "PASS"
exit 0
