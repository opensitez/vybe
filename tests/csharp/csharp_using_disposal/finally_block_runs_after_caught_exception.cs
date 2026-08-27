// vybe-test: csharp/csharp_using_disposal/finally_block_runs_after_caught_exception
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal.rs

using static __Harness;

try { throw new System.Exception(); }
catch (System.Exception) { __P(("caught").ToString()); }
finally { __P(("finally").ToString()); }
__Check("caught\nfinally");

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
