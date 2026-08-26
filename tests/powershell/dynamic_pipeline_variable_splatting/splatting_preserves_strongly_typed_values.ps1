# vybe-test: powershell/dynamic_pipeline_variable_splatting/splatting_preserves_strongly_typed_values
function Target-TypedSplat {
    param([guid]$Id, [datetime]$Date)
    return "ID:$($Id.ToString()),D:$($Date.Year)"
}
$g = [guid]::NewGuid()
$d = [datetime]::Parse("2026-08-26")
$p = @{ Id = $g; Date = $d }
$res = Target-TypedSplat @p
if ($res -ne "ID:$($g.ToString()),D:2026") {
    Write-Host "FAIL: Splatting preserving strongly typed values failed"
    exit 1
}
Write-Host "PASS"
exit 0
