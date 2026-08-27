// vybe-test: csharp/csharp_lambdas/higher_order_function
// origin: languages/csharp/tests/csharp/test_csharp_lambdas.rs

using static __Harness;

Func<int, int> square = x => x * x;
Func<int, int> negate = x => -x;
__P((square(5)).ToString());
__P((negate(5)).ToString());
__Check("25\n-5");

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
