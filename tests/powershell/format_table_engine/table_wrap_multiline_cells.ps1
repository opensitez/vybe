# vybe-test: powershell/format_table_engine/table_wrap_multiline_cells
$data = @([pscustomobject]@{ Item = "MultiLine"; Notes = "FirstLine`nSecondLine" })

# -Wrap preserves line breaks and wraps text across multiple terminal rows
$output = $data | Format-Table -Property Item, Notes -Wrap | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

# Both lines must appear in the rendered output
if (-not ($output -match "FirstLine" -and $output -match "SecondLine")) {
    Write-Host "FAIL: wrapped multi-line cell content missing: $output"
    exit 1
}

Write-Host "PASS"
exit 0
