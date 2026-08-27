// vybe-test: csharp/csharp_lambda_loop_capture_semantics/for_loop_lambda_captures_shared_counter_showing_last_value_at_invoke_time
// origin: languages/csharp/tests/csharp/test_csharp_lambda_loop_capture_semantics.rs

using static __Harness;
using System;
using System.Collections.Generic;

var actions = new List<Func<int>>();
for (int i = 0; i < 3; i++) {
    actions.Add(() => i);
}
foreach (var run in actions) __P((run()).ToString());
__Check("3\n3\n3");

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
