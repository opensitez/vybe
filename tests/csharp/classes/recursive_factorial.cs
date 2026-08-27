// vybe-test: csharp/classes/recursive_factorial
// origin: languages/csharp/tests/csharp/test_classes.rs

using static __Harness;

__P((MathUtils.Fact(5)).ToString());
__Check("120");

class MathUtils {
            public static int Fact(int n) {
                if (n <= 1) return 1;
                return Fact(n - 1) * n;
            }
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
