// vybe-test: csharp/csharp_math_advanced/math_bit_decrement_returns_next_lower_double
// origin: languages/csharp/tests/csharp/test_csharp_math_advanced.rs

using static __Harness;

__P((System.Math.BitDecrement(1.0)<1.0).ToString());
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
