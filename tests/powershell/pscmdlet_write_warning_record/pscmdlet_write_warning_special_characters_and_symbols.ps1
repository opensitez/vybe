# vybe-test: powershell/pscmdlet_write_warning_record/pscmdlet_write_warning_special_characters_and_symbols
# $PSCmdlet.WriteWarning retains special characters, symbols, and formatting unaltered
function EmitSpecialWarning {
    [CmdletBinding()]
    param()
    process {
        $msg = 'SpecialSymbols: [WARN] (CPU > 90%): {user: "admin", code: 0xDEADBEEF}'
        $PSCmdlet.WriteWarning($msg)
    }
}

$oldPreference = $WarningPreference
$WarningPreference = "Continue"

try {
    $record = (EmitSpecialWarning 3>&1)[0]
    $expected = 'SpecialSymbols: [WARN] (CPU > 90%): {user: "admin", code: 0xDEADBEEF}'

    if ($record.Message -ne $expected) {
        Write-Host "FAIL: special symbol warning message mismatch"
        exit 1
    }
} finally {
    $WarningPreference = $oldPreference
}

Write-Host "PASS"
exit 0
