// vybe-test: csharp/csharp_numeric_precision/decimal_avoids_float_rounding_error
// origin: languages/csharp/tests/csharp/test_csharp_numeric_precision.rs

using static __Harness;

decimal a=0.1m, b=0.2m;
__P((a+b==0.3m).ToString());
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
