// vybe-test: csharp/csharp_decimal_financial/decimal_financial_tip_fifteen_percent
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

using static __Harness;

decimal bill = 47.80m;
decimal tip = bill * 0.15m;
__P(tip.ToString("F4", System.Globalization.CultureInfo.InvariantCulture));
__Check("7.1700");
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
