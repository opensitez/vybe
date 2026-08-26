# vybe-test: powershell/json_custom_object_roundtrip/custom_class_instance_roundtrip_to_pscustomobject
class PersonRecordClass {
    [string]$First
    [string]$Last
    PersonRecordClass([string]$f, [string]$l) {
        $this.First = $f
        $this.Last = $l
    }
}
$person = [PersonRecordClass]::new("John", "Doe")
$json = $person | ConvertTo-Json
$recovered = $json | ConvertFrom-Json
if ($recovered.First -ne "John" -or $recovered.Last -ne "Doe") {
    Write-Host "FAIL: Custom class instance roundtrip failed"
    exit 1
}
Write-Host "PASS"
exit 0
