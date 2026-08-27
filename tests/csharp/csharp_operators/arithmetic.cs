// vybe-test: csharp/csharp_operators/arithmetic
// origin: languages/csharp/tests/csharp/test_csharp_operators.rs

using static __Harness;

__P((10 + 5).ToString());
__P((10 - 5).ToString());
__P((10 * 5).ToString());
__P((10 % 3).ToString());
__Check("15\n5\n50\n1");

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
