// vybe-test: csharp/csharp_oop_inheritance/object_tostring_is_overridable_for_custom_display
// origin: languages/csharp/tests/csharp/test_csharp_oop_inheritance.rs

using static __Harness;

__P((new Point { X=1, Y=2 }).ToString());
__Check("(1,2)");

class Point { public int X,Y; public override string ToString() => $"({X},{Y})"; }

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
