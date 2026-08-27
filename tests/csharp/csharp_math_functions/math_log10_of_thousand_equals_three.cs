// vybe-test: csharp/csharp_math_functions/math_log10_of_thousand_equals_three
// origin: languages/csharp/tests/csharp/test_csharp_math_functions.rs

using static __Harness;

__P((System.Math.Log10(1000)).ToString());
__Check("3");

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
