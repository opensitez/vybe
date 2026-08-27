// vybe-test: csharp/csharp_decimal_financial/decimal_financial_compound_two_percent
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

using static __Harness;

decimal principal = 1000.00m;
decimal rate = 0.02m;
decimal total = principal * (1.0m + rate);
__P(total.ToString("F4", System.Globalization.CultureInfo.InvariantCulture));
__Check("1020.0000");
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
