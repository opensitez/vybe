// vybe-test: csharp/csharp_null_propagation/delegate_null_conditional_invocation_skips_when_delegate_is_null
// origin: languages/csharp/tests/csharp/test_csharp_null_propagation.rs

using static __Harness;
using System;

Action action = null;
action?.Invoke();
__P(("done").ToString());
__Check("done");

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
