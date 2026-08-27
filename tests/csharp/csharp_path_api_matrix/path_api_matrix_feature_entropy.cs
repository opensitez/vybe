// vybe-test: csharp/csharp_path_api_matrix/path_api_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_path_api_matrix.rs

using static __Harness;

// path_api_matrix
string feature = "path_api_matrix:123";
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
