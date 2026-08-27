// vybe-test: csharp/csharp_decimal_semantics/decimal_equality_compares_numeric_value_not_reference
// origin: languages/csharp/tests/csharp/test_csharp_decimal_semantics.rs

using static __Harness;

decimal left = 1.0m;
decimal right = 1.00m;
__P((left == right).ToString());
__Check("True");

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
