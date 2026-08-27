// vybe-test: csharp/csharp_math_advanced/math_clamp_returns_value_when_in_range
// origin: languages/csharp/tests/csharp/test_csharp_math_advanced.rs

using static __Harness;

__P((System.Math.Clamp(5,0,10)).ToString());
__Check("5");

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
