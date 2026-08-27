// vybe-test: csharp/csharp_events_advanced/delegate_return_value_can_be_consumed_by_caller
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

using static __Harness;
using System;

__P((Calculator.Compute(() => 9)).ToString());
__Check("10");

class Calculator { public static int Compute(Func<int> getValue) { return getValue() + 1; } }

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
