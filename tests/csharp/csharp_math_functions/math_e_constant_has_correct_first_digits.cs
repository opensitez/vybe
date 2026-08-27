// vybe-test: csharp/csharp_math_functions/math_e_constant_has_correct_first_digits
// origin: languages/csharp/tests/csharp/test_csharp_math_functions.rs

using static __Harness;

__P((System.Math.E > 2.71 && System.Math.E < 2.72).ToString());
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
