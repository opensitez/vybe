// vybe-test: csharp/csharp_with_expression_matrix/with_expression_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_matrix.rs

using static __Harness;

// with_expression_matrix
int? maybe = 108;
__P((maybe.HasValue && maybe.Value == 108).ToString());
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
