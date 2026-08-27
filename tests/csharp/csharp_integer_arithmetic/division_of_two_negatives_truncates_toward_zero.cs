// vybe-test: csharp/csharp_integer_arithmetic/division_of_two_negatives_truncates_toward_zero
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

using static __Harness;

__P((-7 / -3).ToString());
__Check("2");

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
