// vybe-test: csharp/csharp_math_functions/math_atan2_computes_angle_from_y_x_coordinates
// origin: languages/csharp/tests/csharp/test_csharp_math_functions.rs

using static __Harness;

double angle = System.Math.Atan2(1, 1);
__P((System.Math.Round(angle, 4)).ToString());
__Check("0.7854");

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
