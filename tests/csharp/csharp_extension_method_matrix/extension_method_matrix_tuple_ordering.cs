// vybe-test: csharp/csharp_extension_method_matrix/extension_method_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_extension_method_matrix.rs

using static __Harness;

// extension_method_matrix
var tuple = (left: 78, right: 79);
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
