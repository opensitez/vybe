// vybe-test: csharp/csharp_exception_finally_control_flow/finally_in_nested_try_runs_before_outer_catch_handles_exception
// origin: languages/csharp/tests/csharp/test_csharp_exception_finally_control_flow.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

try {
    try {
        throw new Exception("boom");
    } finally {
        __Check(("inner-finally").ToString(), "inner-finally");
    }
} catch (Exception) {
    __Check(("outer-catch").ToString(), "outer-catch");
}
