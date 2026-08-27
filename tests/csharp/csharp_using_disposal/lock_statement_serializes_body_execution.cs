// vybe-test: csharp/csharp_using_disposal/lock_statement_serializes_body_execution
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal.rs

using static __Harness;

object gate = new object();
lock (gate) { __P(("locked").ToString()); }
__Check("locked");

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
