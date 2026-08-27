// vybe-test: csharp/linq_lambdas/func_chain
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

using static __Harness;

Func<int, int> doubleIt = x => x * 2;
Func<int, int> addOne = x => x + 1;
__P((addOne(doubleIt(5))).ToString());
__P((doubleIt(addOne(5))).ToString());
__Check("11\n12");

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
