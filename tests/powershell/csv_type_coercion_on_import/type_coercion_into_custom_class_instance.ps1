# vybe-test: powershell/csv_type_coercion_on_import/type_coercion_into_custom_class_instance
class CsvItemRecord {
    [string]$Name
    [int]$Count
    [bool]$Active
    CsvItemRecord([pscustomobject]$csvRow) {
        $this.Name = $csvRow.Name
        $this.Count = [int]$csvRow.Count
        $this.Active = [bool]::Parse($csvRow.Active)
    }
}
$csv = "Name,Count,Active`nWidget,42,True"
$row = $csv | ConvertFrom-Csv
$record = [CsvItemRecord]::new($row)
if ($record.Count -ne 42 -or $record.Active -ne $true -or $record.Name -ne "Widget") {
    Write-Host "FAIL: Custom class constructor coercion from CSV failed"
    exit 1
}
Write-Host "PASS"
exit 0
