// vybe-test: csharp/csharp_events_advanced/func_delegate_returns_computed_result
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

using static __Harness;
using System;

Func<int, int, int> add = (left, right) => left + right;
__P((add(4, 5)).ToString());
__Check("9");

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
