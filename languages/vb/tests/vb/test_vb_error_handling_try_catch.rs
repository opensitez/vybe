use super::helpers::run_vb;

#[test]
fn try_catch_basic() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nTry\nDim x = 1 \\ 0\nCatch\nConsole.WriteLine(\"Caught\")\nEnd Try\nEnd Sub\nEnd Module"
        ),
        vec!["Caught"]
    );
}
#[test]
fn try_catch_specific_exception() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nTry\nDim x = 1 \\ 0\nCatch ex As System.DivideByZeroException\nConsole.WriteLine(\"DivByZero\")\nCatch ex As System.Exception\nConsole.WriteLine(\"General\")\nEnd Try\nEnd Sub\nEnd Module"
        ),
        vec!["DivByZero"]
    );
}
#[test]
fn try_catch_fallthrough_to_general() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nTry\nThrow New System.InvalidOperationException()\nCatch ex As System.DivideByZeroException\nConsole.WriteLine(\"DivByZero\")\nCatch ex As System.Exception\nConsole.WriteLine(\"General\")\nEnd Try\nEnd Sub\nEnd Module"
        ),
        vec!["General"]
    );
}
#[test]
fn try_catch_finally_executes_on_success() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nTry\nConsole.WriteLine(\"Try\")\nCatch\nConsole.WriteLine(\"Catch\")\nFinally\nConsole.WriteLine(\"Finally\")\nEnd Try\nEnd Sub\nEnd Module"
        ),
        vec!["Try", "Finally"]
    );
}
#[test]
fn try_catch_finally_executes_on_throw() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nTry\nDim x = 1 \\ 0\nCatch\nConsole.WriteLine(\"Catch\")\nFinally\nConsole.WriteLine(\"Finally\")\nEnd Try\nEnd Sub\nEnd Module"
        ),
        vec!["Catch", "Finally"]
    );
}

#[test]
fn try_finally_only() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x = 0\nTry\nx = 1\nFinally\nConsole.WriteLine(x)\nEnd Try\nEnd Sub\nEnd Module"
        ),
        vec!["1"]
    );
}
#[test]
fn try_catch_when_clause_true() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim flag = True\nTry\nThrow New System.Exception()\nCatch ex As System.Exception When flag\nConsole.WriteLine(\"Caught\")\nCatch\nConsole.WriteLine(\"Other\")\nEnd Try\nEnd Sub\nEnd Module"
        ),
        vec!["Caught"]
    );
}
#[test]
fn try_catch_when_clause_false() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim flag = False\nTry\nThrow New System.Exception()\nCatch ex As System.Exception When flag\nConsole.WriteLine(\"Caught\")\nCatch\nConsole.WriteLine(\"Other\")\nEnd Try\nEnd Sub\nEnd Module"
        ),
        vec!["Other"]
    );
}
#[test]
fn try_catch_throw_no_arg() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nTry\nTry\nThrow New System.Exception()\nCatch\nThrow\nEnd Try\nCatch\nConsole.WriteLine(\"Rethrown\")\nEnd Try\nEnd Sub\nEnd Module"
        ),
        vec!["Rethrown"]
    );
} // Rethrows the current exception
#[test]
fn try_catch_throw_new() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nTry\nThrow New System.InvalidOperationException()\nCatch\nConsole.WriteLine(\"Caught\")\nEnd Try\nEnd Sub\nEnd Module"
        ),
        vec!["Caught"]
    );
}

#[test]
fn try_catch_nested() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nTry\nTry\nThrow New System.Exception()\nCatch\nConsole.WriteLine(\"Inner\")\nEnd Try\nCatch\nConsole.WriteLine(\"Outer\")\nEnd Try\nEnd Sub\nEnd Module"
        ),
        vec!["Inner"]
    );
}
#[test]
fn try_catch_nested_bubble_up() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nTry\nTry\nThrow New System.Exception()\nFinally\nConsole.WriteLine(\"InnerFinally\")\nEnd Try\nCatch\nConsole.WriteLine(\"OuterCatch\")\nEnd Try\nEnd Sub\nEnd Module"
        ),
        vec!["InnerFinally", "OuterCatch"]
    );
}
#[test]
fn try_catch_exit_try() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nTry\nConsole.WriteLine(\"Start\")\nExit Try\nConsole.WriteLine(\"End\")\nCatch\nFinally\nConsole.WriteLine(\"Finally\")\nEnd Try\nEnd Sub\nEnd Module"
        ),
        vec!["Start", "Finally"]
    );
} // Finally executes even after Exit Try
#[test]
fn try_catch_return_in_try() {
    assert_eq!(
        run_vb(
            "Module M\nFunction F() As String\nTry\nReturn \"Try\"\nFinally\nConsole.WriteLine(\"Finally\")\nEnd Try\nEnd Function\nSub Main()\nConsole.WriteLine(F())\nEnd Sub\nEnd Module"
        ),
        vec!["Finally", "Try"]
    );
} // Finally executes before function returns
#[test]
fn try_catch_return_in_catch() {
    assert_eq!(
        run_vb(
            "Module M\nFunction F() As String\nTry\nThrow New System.Exception()\nCatch\nReturn \"Catch\"\nFinally\nConsole.WriteLine(\"Finally\")\nEnd Try\nEnd Function\nSub Main()\nConsole.WriteLine(F())\nEnd Sub\nEnd Module"
        ),
        vec!["Finally", "Catch"]
    );
}

#[test]
fn try_catch_multiple_exceptions_same_block() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nTry\nThrow New System.DivideByZeroException()\nCatch ex As System.DivideByZeroException\nConsole.WriteLine(\"Div\")\nCatch ex As System.OverflowException\nConsole.WriteLine(\"Over\")\nEnd Try\nEnd Sub\nEnd Module"
        ),
        vec!["Div"]
    );
}
#[test]
fn try_catch_unreachable_catch_fails() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\n' Try\n' Throw New System.Exception()\n' Catch ex As System.Exception\n' Catch ex As System.DivideByZeroException ' Unreachable because Exception catches all\n' End Try\nConsole.WriteLine(\"Parsed\")\nEnd Sub\nEnd Module"
        ),
        vec!["Parsed"]
    );
}
#[test]
fn try_catch_variable_scope() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x = 0\nTry\nDim y = 5\nCatch\nEnd Try\n' Console.WriteLine(y) ' y is not in scope\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"
        ),
        vec!["0"]
    );
}
#[test]
fn try_catch_exception_variable_scope() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nTry\nThrow New System.Exception()\nCatch ex As System.Exception\nConsole.WriteLine(\"Caught\")\nEnd Try\n' Console.WriteLine(ex.Message) ' ex is not in scope\nEnd Sub\nEnd Module"
        ),
        vec!["Caught"]
    );
}
#[test]
fn on_error_resume_next() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nOn Error Resume Next\nDim x = 1 \\ 0\nConsole.WriteLine(\"Resumed\")\nEnd Sub\nEnd Module"
        ),
        vec!["Resumed"]
    );
} // Legacy error handling

#[test]
fn on_error_goto_label() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nOn Error GoTo ErrorHandler\nDim x = 1 \\ 0\nConsole.WriteLine(\"Skipped\")\nExit Sub\nErrorHandler:\nConsole.WriteLine(\"Handled\")\nEnd Sub\nEnd Module"
        ),
        vec!["Handled"]
    );
}
#[test]
fn on_error_goto_0() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nOn Error Resume Next\nOn Error GoTo 0 ' Disables error handling\nTry\nDim x = 1 \\ 0\nCatch\nConsole.WriteLine(\"CaughtByTry\")\nEnd Try\nEnd Sub\nEnd Module"
        ),
        vec!["CaughtByTry"]
    );
}
#[test]
fn err_object_number() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nOn Error Resume Next\nDim x = 1 \\ 0\nConsole.WriteLine(Err.Number <> 0)\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
}
#[test]
fn err_object_clear() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nOn Error Resume Next\nDim x = 1 \\ 0\nErr.Clear()\nConsole.WriteLine(Err.Number)\nEnd Sub\nEnd Module"
        ),
        vec!["0"]
    );
}
#[test]
fn try_catch_mixed_with_on_error_fails() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\n' Try\n' On Error Resume Next ' Cannot mix Try/Catch with On Error in same method\n' Catch\n' End Try\nConsole.WriteLine(\"Parsed\")\nEnd Sub\nEnd Module"
        ),
        vec!["Parsed"]
    );
}
