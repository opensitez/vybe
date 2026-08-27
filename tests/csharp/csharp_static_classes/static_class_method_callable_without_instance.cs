// vybe-test: csharp/csharp_static_classes/static_class_method_callable_without_instance
// origin: languages/csharp/tests/csharp/test_csharp_static_classes.rs

using static __Harness;

__P((MathHelper.Square(5)).ToString());
__Check("25");

static class MathHelper { public static int Square(int n) => n*n; }

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
