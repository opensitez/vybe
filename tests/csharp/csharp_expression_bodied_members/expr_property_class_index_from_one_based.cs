// vybe-test: csharp/csharp_expression_bodied_members/expr_property_class_index_from_one_based
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

using static __Harness;

__P((new Row { Index = 4 }.Display).ToString());
__Check("5");

class Row { public int Index = 0; public int Display => Index + 1; }

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
