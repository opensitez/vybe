// vybe-test: csharp/csharp_expression_bodied_members/expr_indexer_on_nested_static_data_via_instance
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

using static __Harness;

__P((new Lookup()[2]).ToString());
__Check("7");

class Lookup { int[] table = { 5, 6, 7 }; public int this[int i] => table[i]; }

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
