# vybe-test: powershell/error_record_script_stack_trace/script_stack_trace_contains_function_names
function LevelThree { throw "StackTraceCrash" }
function LevelTwo { LevelThree }
function LevelOne { LevelTwo }
$err = $null
try {
    LevelOne
} catch {
    $err = $_
}
$trace = $err.ScriptStackTrace
if (-not ($trace.Contains("LevelThree") -and $trace.Contains("LevelTwo") -and $trace.Contains("LevelOne"))) {
    Write-Host "FAIL: ScriptStackTrace should contain all call frames, got '$trace'"
    exit 1
}
Write-Host "PASS"
exit 0
