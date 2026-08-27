// vybe-test: csharp/csharp_math_trigonometry_matrix/math_trigonometry_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_math_trigonometry_matrix.rs

using static __Harness;

// math_trigonometry_matrix
var tuple = (left: 102, right: 103);
__P((tuple.left < tuple.right).ToString());
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
