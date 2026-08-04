// vybe-test: csharp/csharp_lambda_loop_capture_semantics/lambda_mutating_captured_local_is_visible_to_later_invocations
// origin: languages/csharp/tests/csharp/test_csharp_lambda_loop_capture_semantics.rs

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

using System;
int tally = 0;
Action bump = () => { tally++; };
bump();
bump();
__P((tally).ToString());
__Check("2");
