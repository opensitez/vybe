// vybe-test: csharp/csharp_exception_finally_control_flow/finally_that_throws_during_return_propagates_past_the_try
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

void M() {
    try {
        try {
            __P(("body").ToString());
            return;
        } finally {
            __P(("finally").ToString());
            throw new Exception("boom");
        }
    } catch (Exception) {
        __P(("caught").ToString());
    }
}
M();
__Check("body\nfinally\ncaught");
