// vybe-test: csharp/csharp_classes/class_static_method
// origin: languages/csharp/tests/csharp/test_csharp_classes.rs

using static __Harness;

__P((MathUtils.Square(7)).ToString());
__Check("49");

class MathUtils {
    public static int Square(int x) { return x * x; }
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
