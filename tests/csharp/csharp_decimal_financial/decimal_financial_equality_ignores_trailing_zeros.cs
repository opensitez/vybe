// vybe-test: csharp/csharp_decimal_financial/decimal_financial_equality_ignores_trailing_zeros
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

using static __Harness;

__P((2.50m==2.5m).ToString());
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
