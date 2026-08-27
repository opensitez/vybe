// vybe-test: csharp/csharp_numeric_checked_bitwise/checked_block_throws_on_overflow_for_byte_addition
// origin: languages/csharp/tests/csharp/test_csharp_numeric_checked_bitwise.rs

using static __Harness;

try { checked { byte value = 255; value += 1; } __P(("no-throw").ToString()); }
catch (System.OverflowException) { __P(("overflow").ToString()); }
__Check("overflow");

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
