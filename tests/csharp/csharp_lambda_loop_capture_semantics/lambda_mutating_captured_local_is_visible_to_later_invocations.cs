// vybe-test: csharp/csharp_lambda_loop_capture_semantics/lambda_mutating_captured_local_is_visible_to_later_invocations
// origin: languages/csharp/tests/csharp/test_csharp_lambda_loop_capture_semantics.rs

using static __Harness;
using System;

int tally = 0;
Action bump = () => { tally++; }
;
bump();
bump();
__P((tally).ToString());
__Check("2");

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
