// vybe-test: csharp/classes/static_method_in_class
// origin: languages/csharp/tests/csharp/test_classes.rs

using static __Harness;

__P((MathUtils.Add(3, 4)).ToString());
__Check("7");

class MathUtils {
            public static int Add(int a, int b) { return a + b; }
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
