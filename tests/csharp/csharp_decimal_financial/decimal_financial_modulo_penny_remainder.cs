// vybe-test: csharp/csharp_decimal_financial/decimal_financial_modulo_penny_remainder
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

using static __Harness;

__P((10.01m%0.10m).ToString());
__Check("0.01");

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
