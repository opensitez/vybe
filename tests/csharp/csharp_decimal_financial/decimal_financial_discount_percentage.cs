// vybe-test: csharp/csharp_decimal_financial/decimal_financial_discount_percentage
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

using static __Harness;

decimal price = 250.00m;
decimal discount = 0.20m;
decimal finalPrice = price * (1.0m - discount);
__P(finalPrice.ToString("F4", System.Globalization.CultureInfo.InvariantCulture));
__Check("200.0000");
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
