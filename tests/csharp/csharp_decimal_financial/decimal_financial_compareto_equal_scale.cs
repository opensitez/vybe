// vybe-test: csharp/csharp_decimal_financial/decimal_financial_compareto_equal_scale
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

using static __Harness;

__P((1.0m.CompareTo(1.00m)).ToString());
__Check("0");

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
