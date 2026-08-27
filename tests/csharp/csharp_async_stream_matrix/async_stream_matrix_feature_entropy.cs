// vybe-test: csharp/csharp_async_stream_matrix/async_stream_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_async_stream_matrix.rs

using static __Harness;

// async_stream_matrix
string feature = "async_stream_matrix:111";
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
