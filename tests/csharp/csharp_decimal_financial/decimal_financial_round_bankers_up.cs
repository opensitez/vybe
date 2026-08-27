// vybe-test: csharp/csharp_decimal_financial/decimal_financial_round_bankers_up
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

using static __Harness;

__P((decimal.Round(2.225m,2)).ToString());
__Check("2.22");

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
