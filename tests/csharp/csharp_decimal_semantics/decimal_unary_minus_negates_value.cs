// vybe-test: csharp/csharp_decimal_semantics/decimal_unary_minus_negates_value
// origin: languages/csharp/tests/csharp/test_csharp_decimal_semantics.rs

using static __Harness;

decimal balance = 12.5m;
__P((-balance).ToString());
__Check("-12.5");

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
