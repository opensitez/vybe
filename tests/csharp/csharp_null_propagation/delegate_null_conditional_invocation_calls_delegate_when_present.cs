// vybe-test: csharp/csharp_null_propagation/delegate_null_conditional_invocation_calls_delegate_when_present
// origin: languages/csharp/tests/csharp/test_csharp_null_propagation.rs

using static __Harness;
using System;

Action action = () => __P(("ran").ToString());
action?.Invoke();
__Check("ran");

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
