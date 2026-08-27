// vybe-test: csharp/classes/static_class_method_call
// origin: languages/csharp/tests/csharp/test_classes.rs

using static __Harness;

__P((MathHelper.Square(5)).ToString());
__P((MathHelper.Double(7)).ToString());
__Check("25\n14");

class MathHelper {
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
