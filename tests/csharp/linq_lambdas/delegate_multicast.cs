// vybe-test: csharp/linq_lambdas/delegate_multicast
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

using static __Harness;

Action<string> logger = msg => __P(("LOG: " + msg).ToString());
Action<string> printer = msg => __P(("PRINT: " + msg).ToString());
Action<string> both = logger + printer;
both("hello");
__Check("LOG: hello\nPRINT: hello");

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
