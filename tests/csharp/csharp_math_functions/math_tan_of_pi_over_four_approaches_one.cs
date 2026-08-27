// vybe-test: csharp/csharp_math_functions/math_tan_of_pi_over_four_approaches_one
// origin: languages/csharp/tests/csharp/test_csharp_math_functions.rs

using static __Harness;

double v = System.Math.Tan(System.Math.PI / 4);
__P((System.Math.Round(v)).ToString());
__Check("1");

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
