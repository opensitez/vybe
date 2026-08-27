// vybe-test: csharp/csharp_numeric_ops/integer_division_by_zero_throws
// origin: languages/csharp/tests/csharp/test_csharp_numeric_ops.rs

using static __Harness;

bool threw = false;
try {
    int zero = 0;
    int x = 10 / zero;
} catch (DivideByZeroException) {
    threw = true;
}
__P(threw.ToString());
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
