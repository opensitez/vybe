// vybe-test: csharp/oop_advanced/static_class_methods
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

using static __Harness;

__P((MathHelper.Square(5)).ToString());
__P((MathHelper.Double(7)).ToString());
__Check("25\n14");

static class MathHelper {
    public static int Square(int x) { return x * x; }
    public static int Double(int x) { return x * 2; }
}

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
