// vybe-test: csharp/csharp_decimal_financial/decimal_financial_tax_rate_application
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

using static __Harness;

decimal amount = 100.00m;
decimal rate = 0.0825m;
decimal tax = amount * rate;
__P(tax.ToString("F6", System.Globalization.CultureInfo.InvariantCulture));
__Check("8.250000");
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
