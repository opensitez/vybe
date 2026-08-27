// vybe-test: csharp/csharp_span_indexing/readonly_span_from_string_has_correct_length
// origin: languages/csharp/tests/csharp/test_csharp_span_indexing.rs

using static __Harness;

System.ReadOnlySpan<char> span = "hello".AsSpan();
__P((span.Length).ToString());
__Check("5");

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
