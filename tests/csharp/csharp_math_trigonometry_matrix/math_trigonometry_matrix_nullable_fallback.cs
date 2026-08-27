// vybe-test: csharp/csharp_math_trigonometry_matrix/math_trigonometry_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_math_trigonometry_matrix.rs

using static __Harness;

// math_trigonometry_matrix
int? maybe = null;
int fallback = maybe ?? 102;
__P((fallback == 102).ToString());
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
