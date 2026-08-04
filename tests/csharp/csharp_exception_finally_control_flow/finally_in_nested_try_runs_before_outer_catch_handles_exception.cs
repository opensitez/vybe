// vybe-test: csharp/csharp_exception_finally_control_flow/finally_in_nested_try_runs_before_outer_catch_handles_exception
// origin: languages/csharp/tests/csharp/test_csharp_exception_finally_control_flow.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

try {
    try {
        throw new Exception("boom");
    } finally {
        __P(("inner-finally").ToString());
    }
} catch (Exception) {
    __P(("outer-catch").ToString());
}
__Check("inner-finally\nouter-catch");
