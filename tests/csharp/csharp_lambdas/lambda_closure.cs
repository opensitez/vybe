// vybe-test: csharp/csharp_lambdas/lambda_closure
// origin: languages/csharp/tests/csharp/test_csharp_lambdas.rs

using static __Harness;

int counter = 0;
var inc = () => { counter++; }
;
inc();
inc();
inc();
__P((counter).ToString());
__Check("3");

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
