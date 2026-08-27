// vybe-test: csharp/csharp_operators/comparison
// origin: languages/csharp/tests/csharp/test_csharp_operators.rs

using static __Harness;

__P((1 < 2).ToString());
__P((2 > 1).ToString());
__P((1 <= 1).ToString());
__P((1 >= 1).ToString());
__P((1 == 1).ToString());
__P((1 != 2).ToString());
__Check("True\nTrue\nTrue\nTrue\nTrue\nTrue");

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
