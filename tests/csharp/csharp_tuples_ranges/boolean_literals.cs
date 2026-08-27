// vybe-test: csharp/csharp_tuples_ranges/boolean_literals
// origin: languages/csharp/tests/csharp/test_csharp_tuples_ranges.rs

using static __Harness;

__P((true).ToString());
__P((false).ToString());
__P((true && false).ToString());
__P((true || false).ToString());
__Check("True\nFalse\nFalse\nTrue");

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
