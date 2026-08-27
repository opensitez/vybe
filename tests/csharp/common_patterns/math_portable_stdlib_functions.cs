// vybe-test: csharp/common_patterns/math_portable_stdlib_functions
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

using static __Harness;

__P((Math.Sin(0)).ToString());
__P((Math.Cos(0)).ToString());
__P((Math.Tan(0)).ToString());
__P((Math.Asin(0)).ToString());
__P((Math.Acos(1)).ToString());
__P((Math.Atan(0)).ToString());
__P((Math.Atan2(0, 1)).ToString());
__P((Math.Log(1)).ToString());
__P((Math.Log10(100)).ToString());
__P((Math.Exp(0)).ToString());
__P((Math.Sign(-5)).ToString());
__P((Math.Clamp(15, 0, 10)).ToString());
__Check("0\n1\n0\n0\n0\n0\n0\n0\n2\n1\n-1\n10");

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
