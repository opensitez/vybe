# vybe-test: powershell/pscmdlet_write_information_record/pscmdlet_write_information_message_data_custom_object
# InformationRecord.MessageData can store structured PSCustomObject payloads without stringification
function EmitObjectInfo {
    [CmdletBinding()]
    param()
    process {
        $data = [PSCustomObject]@{
            HostName = "app-worker-1"
            Port = 8080
            Healthy = $true
        }
        $PSCmdlet.WriteInformation($data, @("Telemetry"))
    }
}

$oldPreference = $InformationPreference
$InformationPreference = "Continue"

try {
    $record = (EmitObjectInfo 6>&1)[0]
    $obj = $record.MessageData

    if ($obj.HostName -ne "app-worker-1" -or $obj.Port -ne 8080 -or ($obj.Healthy -ne $true)) {
        Write-Host "FAIL: structured object property mismatch in MessageData"
        exit 1
    }
} finally {
    $InformationPreference = $oldPreference
}

Write-Host "PASS"
exit 0
