// vybe-test: csharp/csharp_expression_bodied_members/expr_method_nested_class_delegates_to_outer
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

using static __Harness;

__P((new Outer.Inner(new Outer()).Boost()).ToString());
__Check("15");

class Outer { public int Base => 10; public class Inner { Outer o; public Inner(Outer owner) { o = owner; } public int Boost() => o.Base + 5; } }

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
