// vybe-test: csharp/csharp_null_conditional_matrix/null_conditional_matrix_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_null_conditional_matrix.rs

using static __Harness;

// null_conditional_matrix
decimal amount = 10m;
__P(((amount / 2m) * 2m == 10m).ToString());
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
