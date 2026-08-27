// vybe-test: csharp/csharp_expression_bodied_members/expr_property_class_get_set_both_expression
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

using static __Harness;

var b = new Box();
b.Value = 9;
__P((b.Value).ToString());
__Check("9");

class Box { int _v; public int Value { get => _v; set => _v = value; } }

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
