# vybe-test: powershell/pipeline_begin_process_end_blocks/process_block_with_custom_objects
function Enhance-Object {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true)][pscustomobject]$Obj)
    process {
        $Obj | Add-Member -NotePropertyName "Processed" -NotePropertyValue $true -PassThru
    }
}
$objs = @([pscustomobject]@{ Id = 1 }, [pscustomobject]@{ Id = 2 })
$res = @($objs | Enhance-Object)
if ($res[0].Processed -ne $true -or $res[1].Processed -ne $true) {
    Write-Host "FAIL: Custom object pipeline enhancement failed"
    exit 1
}
Write-Host "PASS"
exit 0
