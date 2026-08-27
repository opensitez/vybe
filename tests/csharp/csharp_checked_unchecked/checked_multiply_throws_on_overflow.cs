// vybe-test: csharp/csharp_checked_unchecked/checked_multiply_throws_on_overflow
// origin: languages/csharp/tests/csharp/test_csharp_checked_unchecked.rs

using static __Harness;

__P("Valid_checked_multiply_throws_on_overflow");
__Check("Valid_checked_multiply_throws_on_overflow");
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
