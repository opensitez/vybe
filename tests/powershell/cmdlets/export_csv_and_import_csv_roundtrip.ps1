# vybe-test: powershell/cmdlets/export_csv_and_import_csv_roundtrip
# Export-Csv serializes custom objects to disk and Import-Csv accurately parses them back
$tmpFile = [System.IO.Path]::GetTempFileName()

try {
    $records = @(
        [PSCustomObject]@{ Service = "ApiGateway"; Port = "443"; Status = "Active" },
        [PSCustomObject]@{ Service = "WorkerQueue"; Port = "5672"; Status = "Idle" }
    )

    $records | Export-Csv -Path $tmpFile -NoTypeInformation

    $imported = @(Import-Csv -Path $tmpFile)

    if ($imported.Count -ne 2) {
        Write-Host "FAIL: expected 2 imported records, got $($imported.Count)"
        exit 1
    }

    if ($imported[0].Service -ne "ApiGateway" -or $imported[0].Port -ne "443") {
        Write-Host "FAIL: first record field mismatch"
        exit 1
    }

    if ($imported[1].Service -ne "WorkerQueue" -or $imported[1].Status -ne "Idle") {
        Write-Host "FAIL: second record field mismatch"
        exit 1
    }
} finally {
    if (Test-Path $tmpFile) {
        Remove-Item $tmpFile -Force
    }
}

Write-Host "PASS"
exit 0
