// vybe-test: csharp/csharp_with_expression_records/with_inline_in_expression
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

using static __Harness;

__P(((new V(2) with{N=7}).N).ToString());
__Check("7");

record V(int N);

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
