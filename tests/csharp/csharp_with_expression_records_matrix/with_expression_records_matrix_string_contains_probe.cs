// vybe-test: csharp/csharp_with_expression_records_matrix/with_expression_records_matrix_string_contains_probe
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records_matrix.rs

using static __Harness;

// with_expression_records_matrix
string feature = "with_expression_records_matrix";
__P((feature.Contains("a") || !feature.Contains("a")).ToString());
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
