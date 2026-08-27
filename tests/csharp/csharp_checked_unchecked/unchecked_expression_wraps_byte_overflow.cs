// vybe-test: csharp/csharp_checked_unchecked/unchecked_expression_wraps_byte_overflow
// origin: languages/csharp/tests/csharp/test_csharp_checked_unchecked.rs

using static __Harness;

byte b=unchecked((byte)256);
__P((b).ToString());
__Check("0");

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
