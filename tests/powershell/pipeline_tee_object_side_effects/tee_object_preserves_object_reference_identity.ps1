# vybe-test: powershell/pipeline_tee_object_side_effects/tee_object_preserves_object_reference_identity
class IdentBox { [int]$Id; IdentBox([int]$i) { $this.Id = $i } }
$box = [IdentBox]::new(99)
$sideBox = $null
$outBox = $box | Tee-Object -Variable sideBox
if ($outBox -ne $box -or $sideBox -ne $box) {
    Write-Host "FAIL: Tee-Object reference identity check failed"
    exit 1
}
Write-Host "PASS"
exit 0
