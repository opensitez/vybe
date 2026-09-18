# vybe-test: powershell/pscmdlet_write_debug_record/pscmdlet_write_debug_special_characters_and_formatting
# $PSCmdlet.WriteDebug preserves special characters, symbols, and formatting verbatim
function EmitSpecialDebug {
    [CmdletBinding()]
    param()
    process {
        $special = 'Symbols: `~!@#$%^&*()_+-=[]{}|\;'':",./<>?'
        $PSCmdlet.WriteDebug($special)
    }
}

$oldPreference = $DebugPreference
$DebugPreference = "Continue"

try {
    $record = (EmitSpecialDebug *>&1)[0]
    $expected = 'Symbols: `~!@#$%^&*()_+-=[]{}|\;'':",./<>?'

    if ($record.Message -ne $expected) {
        Write-Host "FAIL: special characters mismatch in debug record"
        exit 1
    }
} finally {
    $DebugPreference = $oldPreference
}

Write-Host "PASS"
exit 0
