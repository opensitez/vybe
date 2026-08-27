// vybe-test: csharp/csharp_decimal_semantics/decimal_modulo_returns_remainder_for_non_integer_division
// origin: languages/csharp/tests/csharp/test_csharp_decimal_semantics.rs

using static __Harness;

__P((10.5m % 3m).ToString());
__Check("1.5");

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
