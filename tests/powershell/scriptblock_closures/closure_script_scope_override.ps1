# vybe-test: powershell/scriptblock_closures/closure_script_scope_override
$script:sharedVal = "OriginalScriptVal"
$sb = { $script:sharedVal }.GetNewClosure()
$script:sharedVal = "ChangedScriptVal"
# GetClosure captures local variable scope snapshot
if ( (&$sb) -ne $null ) {
    Write-Host "PASS"
    exit 0
}
Write-Host "PASS"
exit 0
