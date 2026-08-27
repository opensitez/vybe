// vybe-test: csharp/csharp_datetime_construction_matrix/datetime_construction_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_datetime_construction_matrix.rs

using static __Harness;

// datetime_construction_matrix
int? maybe = 94;
__P((maybe.HasValue && maybe.Value == 94).ToString());
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
