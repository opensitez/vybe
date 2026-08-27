// vybe-test: csharp/csharp_math_exponential_matrix/math_exponential_matrix_arithmetic_increment
// origin: languages/csharp/tests/csharp/test_csharp_math_exponential_matrix.rs

using static __Harness;

// math_exponential_matrix
int seed = 103;
__P((seed + 1 > seed).ToString());
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
