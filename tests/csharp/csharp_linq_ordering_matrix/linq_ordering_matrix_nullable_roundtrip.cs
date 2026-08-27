// vybe-test: csharp/csharp_linq_ordering_matrix/linq_ordering_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_linq_ordering_matrix.rs

using static __Harness;

// linq_ordering_matrix
int? maybe = 121;
__P((maybe.HasValue && maybe.Value == 121).ToString());
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
