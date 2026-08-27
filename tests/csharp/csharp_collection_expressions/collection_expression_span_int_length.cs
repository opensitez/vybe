// vybe-test: csharp/csharp_collection_expressions/collection_expression_span_int_length
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

using static __Harness;

System.Span<int> s = [1, 2, 3];
__P((s.Length).ToString());
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
