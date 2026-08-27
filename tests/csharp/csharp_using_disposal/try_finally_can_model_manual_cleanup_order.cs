// vybe-test: csharp/csharp_using_disposal/try_finally_can_model_manual_cleanup_order
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal.rs

using static __Harness;

try { __P(("body").ToString()); }
finally { __P(("cleanup").ToString()); }
__Check("body\ncleanup");

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
