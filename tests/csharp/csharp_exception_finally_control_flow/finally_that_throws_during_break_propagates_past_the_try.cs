// vybe-test: csharp/csharp_exception_finally_control_flow/finally_that_throws_during_break_propagates_past_the_try
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

string trace = "";
try {
    for (int i = 0; i < 3; i++) {
        try {
            trace += "body;";
            break;
        } finally {
            trace += "finally;";
            throw new Exception("boom");
        }
    }
    trace += "unreachable;";
} catch (Exception) {
    trace += "caught";
}
__P((trace).ToString());
__Check("body;finally;caught");
