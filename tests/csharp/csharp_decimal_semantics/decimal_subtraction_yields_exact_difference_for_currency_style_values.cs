// vybe-test: csharp/csharp_decimal_semantics/decimal_subtraction_yields_exact_difference_for_currency_style_values
// origin: languages/csharp/tests/csharp/test_csharp_decimal_semantics.rs

using static __Harness;

decimal price = 19.99m;
decimal discount = 4.50m;
__P((price - discount).ToString());
__Check("15.49");

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
