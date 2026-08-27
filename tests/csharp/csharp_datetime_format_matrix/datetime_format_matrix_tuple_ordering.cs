// vybe-test: csharp/csharp_datetime_format_matrix/datetime_format_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_datetime_format_matrix.rs

using static __Harness;

// datetime_format_matrix
var tuple = (left: 96, right: 97);
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
