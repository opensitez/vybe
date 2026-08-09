# vybe-test: powershell/should_process/should_process_in_module
$m = New-Module -ScriptBlock {
    function Exported-ShouldProc {
        [CmdletBinding(SupportsShouldProcess=$true)]
        param([string]$Key)
        if ($PSCmdlet.ShouldProcess($Key)) { return "ModProcOK" }
    }
}
$res = & $m { Exported-ShouldProc "ModKey" }
if ($res -ne "ModProcOK") {
    Write-Host "FAIL: module function ShouldProcess expected ModProcOK, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
