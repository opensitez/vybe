// vybe-test: csharp/csharp_numeric_checked_bitwise/decimal_addition_preserves_decimal_precision
// origin: languages/csharp/tests/csharp/test_csharp_numeric_checked_bitwise.rs

using static __Harness;

decimal left = 1.2m;
decimal right = 2.3m;
__P((left + right).ToString());
__Check("3.5");

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
