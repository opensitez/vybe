// vybe-test: csharp/linq_lambdas/lambda_passed_to_method
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

using static __Harness;

var p = new Processor();
__P((p.Apply(5, x => x * x)).ToString());
__P((p.Apply(5, x => x + 10)).ToString());
__Check("25\n15");

class Processor {
    public int Apply(int value, Func<int, int> transform) {
        return transform(value);
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
