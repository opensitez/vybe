// vybe-test: csharp/csharp_conditional_expressions/null_coalescing_returns_left_when_non_null
// origin: languages/csharp/tests/csharp/test_csharp_conditional_expressions.rs

using static __Harness;

string s="hello";
__P((s??"default").ToString());
__Check("hello");

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
