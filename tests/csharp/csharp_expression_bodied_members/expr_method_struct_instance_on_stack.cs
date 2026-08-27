// vybe-test: csharp/csharp_expression_bodied_members/expr_method_struct_instance_on_stack
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

using static __Harness;

var c = new Counter();
__P((c.Next()).ToString());
__P((c.Next()).ToString());
__Check("1\n2");

struct Counter { public int n; public int Next() => ++n; }

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
