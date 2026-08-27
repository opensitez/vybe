// vybe-test: csharp/csharp_integer_arithmetic/mixed_expression_with_pre_decrement_and_multiplication
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

using static __Harness;

int value = 6;
__P((--value * 2).ToString());
__Check("10");

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
