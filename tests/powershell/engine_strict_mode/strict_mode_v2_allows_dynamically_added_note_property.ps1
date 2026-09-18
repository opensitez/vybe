# vybe-test: powershell/engine_strict_mode/strict_mode_v2_allows_dynamically_added_note_property
Set-StrictMode -Version 2.0

$obj = [pscustomobject]@{ BaseProp = "init" }

# In v2.0, dynamically attached NoteProperties via Add-Member become valid properties that can be read without error
$obj | Add-Member -NotePropertyName "DynamicProp" -NotePropertyValue 999

$caught = $false
$val = $null
try {
    $val = $obj.DynamicProp
} catch {
    $caught = $true
}

if ($caught) {
    Write-Host "FAIL: accessing dynamically added NoteProperty threw under v2.0"
    exit 1
}

if ($val -ne 999) {
    Write-Host "FAIL: expected DynamicProp to be 999, got $val"
    exit 1
}

Write-Host "PASS"
exit 0
