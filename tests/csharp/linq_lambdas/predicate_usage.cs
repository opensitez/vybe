// vybe-test: csharp/linq_lambdas/predicate_usage
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

using static __Harness;

Predicate<int> isEven = x => x % 2 == 0;
__P((isEven(4)).ToString());
__P((isEven(7)).ToString());
__Check("True\nFalse");

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
