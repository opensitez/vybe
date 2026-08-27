// vybe-test: csharp/csharp_collection_expressions/collection_expression_short_array
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

using static __Harness;

short[] vals = [10, 20];
__P((vals[1]).ToString());
__Check("20");

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
