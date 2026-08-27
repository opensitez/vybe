// vybe-test: csharp/csharp_numeric_precision/double_has_floating_point_rounding
// origin: languages/csharp/tests/csharp/test_csharp_numeric_precision.rs

using static __Harness;

double a=0.1, b=0.2;
__P((a+b==0.3).ToString());
__Check("False");

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
