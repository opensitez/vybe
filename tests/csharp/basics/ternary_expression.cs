// vybe-test: csharp/basics/ternary_expression
// origin: languages/csharp/tests/csharp/test_basics.rs

using static __Harness;

__P((5 > 3 ? "yes" : "no").ToString());
__Check("yes");

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
