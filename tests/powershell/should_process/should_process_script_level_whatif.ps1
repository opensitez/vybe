# vybe-test: powershell/should_process/should_process_script_level_whatif
[CmdletBinding(SupportsShouldProcess=$true)]
param([string]$Arg = "Default")

if ($PSCmdlet.ShouldProcess($Arg, "TestScriptWhatIf")) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL: script level ShouldProcess evaluation failed"
exit 1
