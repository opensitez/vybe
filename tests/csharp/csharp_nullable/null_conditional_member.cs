// vybe-test: csharp/csharp_nullable/null_conditional_member
// origin: languages/csharp/tests/csharp/test_csharp_nullable.rs

using static __Harness;

Wrapper w = null;
__P((w?.Value ?? "null").ToString());
w = new Wrapper("hello");
__P((w?.Value ?? "null").ToString());
__Check("null\nhello");

class Wrapper {
    public string Value;
    public Wrapper(string v) { Value = v; }
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
