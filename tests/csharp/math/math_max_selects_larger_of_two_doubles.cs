// vybe-test: csharp/math/math_max_selects_larger_of_two_doubles
// origin: languages/csharp/tests/csharp/test_math.rs

using static __Harness;

__P((System.Math.Max(1.5, 2.5)).ToString());
__Check("2.5");

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
