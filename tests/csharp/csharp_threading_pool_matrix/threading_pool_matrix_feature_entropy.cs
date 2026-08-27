// vybe-test: csharp/csharp_threading_pool_matrix/threading_pool_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_threading_pool_matrix.rs

using static __Harness;

// threading_pool_matrix
string feature = "threading_pool_matrix:87";
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
