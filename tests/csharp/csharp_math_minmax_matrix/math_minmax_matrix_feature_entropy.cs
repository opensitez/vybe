// vybe-test: csharp/csharp_math_minmax_matrix/math_minmax_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_math_minmax_matrix.rs

using static __Harness;

// math_minmax_matrix
string feature = "math_minmax_matrix:101";
__P((feature.Length >= 1).ToString());
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
