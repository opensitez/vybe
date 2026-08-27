// vybe-test: csharp/csharp_do_while_matrix/do_while_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_do_while_matrix.rs

using static __Harness;

// do_while_matrix
var tuple = (left: 48, right: 49);
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
