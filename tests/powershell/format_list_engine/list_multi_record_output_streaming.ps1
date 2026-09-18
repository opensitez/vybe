# vybe-test: powershell/format_list_engine/list_multi_record_output_streaming
$records = @(
    [pscustomobject]@{ RecordId = "REC-1"; Owner = "Alice" },
    [pscustomobject]@{ RecordId = "REC-2"; Owner = "Bob" }
)

# Piping multiple records into Format-List produces distinct list blocks separated by blank lines
$output = $records | Format-List -Property RecordId, Owner | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

if (-not ($output -match "RecordId\s*:\s*REC-1" -and 
          $output -match "Owner\s*:\s*Alice" -and 
          $output -match "RecordId\s*:\s*REC-2" -and 
          $output -match "Owner\s*:\s*Bob")) {
    Write-Host "FAIL: multi-record streaming list failed: $output"
    exit 1
}

Write-Host "PASS"
exit 0
