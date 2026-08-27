// vybe-test: csharp/csharp_lambda_loop_capture_semantics/foreach_iteration_lambda_sees_each_elements_value_not_final_index
// origin: languages/csharp/tests/csharp/test_csharp_lambda_loop_capture_semantics.rs

using static __Harness;
using System;
using System.Collections.Generic;

var actions = new List<Func<int>>();
foreach (var value in new[] { 10, 20, 30 }) {
    actions.Add(() => value);
}
foreach (var run in actions) __P((run()).ToString());
__Check("10\n20\n30");

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
