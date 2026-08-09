# vybe-test: powershell/psnote_properties/psnote_property_pipeline_input
$res = 1..2 | ForEach-Object {
    $o = [pscustomobject]@{ Num = $_ }
    $o | Add-Member -NotePropertyName "Double" -NotePropertyValue ($_ * 2) -PassThru
}
if ($res[0].Double -ne 2 -or $res[1].Double -ne 4) {
    Write-Host "FAIL: pipeline Add-Member NoteProperty expected Double 2, 4"
    exit 1
}
Write-Host "PASS"
exit 0
