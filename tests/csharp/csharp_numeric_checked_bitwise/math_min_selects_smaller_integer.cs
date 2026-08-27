// vybe-test: csharp/csharp_numeric_checked_bitwise/math_min_selects_smaller_integer
// origin: languages/csharp/tests/csharp/test_csharp_numeric_checked_bitwise.rs

using static __Harness;

__P((System.Math.Min(4, 7)).ToString());
__Check("4");

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
