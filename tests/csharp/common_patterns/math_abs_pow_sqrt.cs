// vybe-test: csharp/common_patterns/math_abs_pow_sqrt
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

using static __Harness;

__P((Math.Abs(-42)).ToString());
__P((Math.Pow(2, 10)).ToString());
__P((Math.Sqrt(144)).ToString());
__Check("42\n1024\n12");

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
