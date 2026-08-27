// vybe-test: csharp/csharp_numeric_ops/double_division_by_zero_produces_infinity
// origin: languages/csharp/tests/csharp/test_csharp_numeric_ops.rs

using static __Harness;

double d=1.0/0.0;
__P((double.IsInfinity(d)).ToString());
__Check("True");

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
