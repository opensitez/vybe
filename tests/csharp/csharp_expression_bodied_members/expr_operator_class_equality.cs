// vybe-test: csharp/csharp_expression_bodied_members/expr_operator_class_equality
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

using static __Harness;

__P((new Tag { Name = "x" } == new Tag { Name = "x" }).ToString());
__P((new Tag { Name = "a" } != new Tag { Name = "b" }).ToString());
__Check("True\nTrue");

class Tag { public string Name; public static bool operator ==(Tag a, Tag b) => a.Name == b.Name; public static bool operator !=(Tag a, Tag b) => !(a == b); }

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
