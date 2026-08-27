// vybe-test: csharp/csharp_lambda_loop_capture_semantics/lambda_passed_to_method_receives_current_binding_not_snapshot_at_creation
// origin: languages/csharp/tests/csharp/test_csharp_lambda_loop_capture_semantics.rs

using static __Harness;
using System;

int total = 1;
Action add = () => total += 4;
void Apply(Action work) { work(); }
Apply(add);
__P((total).ToString());
__Check("5");

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
