// vybe-test: csharp/csharp_string_span/readonly_span_index_from_end_works
// origin: languages/csharp/tests/csharp/test_csharp_string_span.rs

using static __Harness;

System.ReadOnlySpan<char> s="hello".AsSpan();
__P((s[^1]).ToString());
__Check("o");

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
