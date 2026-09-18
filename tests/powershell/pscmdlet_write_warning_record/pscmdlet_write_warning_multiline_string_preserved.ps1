# vybe-test: powershell/pscmdlet_write_warning_record/pscmdlet_write_warning_multiline_string_preserved
# Multiline string messages in $PSCmdlet.WriteWarning retain all embedded newlines in WarningRecord.Message
function EmitMultilineWarning {
    [CmdletBinding()]
    param()
    process {
        $text = "Warning Level: Moderate`nService: PaymentGateway`nAction Required: Inspect certificates"
        $PSCmdlet.WriteWarning($text)
    }
}

$oldPreference = $WarningPreference
$WarningPreference = "Continue"

try {
    $record = (EmitMultilineWarning 3>&1)[0]

    $lines = $record.Message -split "`n"
    if ($lines.Count -ne 3) {
        Write-Host "FAIL: expected 3 lines in multiline warning, got $($lines.Count)"
        exit 1
    }

    if ($lines[0] -ne "Warning Level: Moderate" -or $lines[2] -ne "Action Required: Inspect certificates") {
        Write-Host "FAIL: multiline warning content mismatch"
        exit 1
    }
} finally {
    $WarningPreference = $oldPreference
}

Write-Host "PASS"
exit 0
