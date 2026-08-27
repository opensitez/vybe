// vybe-test: csharp/csharp_decimal_semantics/decimal_comparison_orders_values_before_string_conversion
// origin: languages/csharp/tests/csharp/test_csharp_decimal_semantics.rs

using static __Harness;

decimal low = 1.2m;
decimal high = 1.3m;
__P((low < high).ToString());
__P((high > low).ToString());
__Check("True\nTrue");

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
