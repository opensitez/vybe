// vybe-test: csharp/csharp_expression_bodied_members/expr_indexer_class_get_set_int_key
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

using static __Harness;

var b = new Buffer();
b[2] = 99;
__P((b[2]).ToString());
__Check("99");

class Buffer { int[] data = new int[3]; public int this[int i] { get => data[i]; set => data[i] = value; } }

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
