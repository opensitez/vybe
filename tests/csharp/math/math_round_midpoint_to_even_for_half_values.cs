// vybe-test: csharp/math/math_round_midpoint_to_even_for_half_values
// origin: languages/csharp/tests/csharp/test_math.rs

using static __Harness;

__P((System.Math.Round(2.5)).ToString());
__Check("2");

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
