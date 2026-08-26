# vybe-test: powershell/error_record_script_stack_trace/script_stack_trace_format_multiline
function F1 { F2 }
function F2 { throw "MultiLineTrace" }
$err = $null
try { F1 } catch { $err = $_ }
$lines = @($err.ScriptStackTrace -split "`r?`n" | Where-Object { $_.Trim() -ne "" })
if ($lines.Length -lt 2) {
    Write-Host "FAIL: Multiline stack trace expected >= 2 lines, got $($lines.Length)"
    exit 1
}
Write-Host "PASS"
exit 0
