// vybe-test: csharp/csharp_string_span/readonly_span_char_from_string_slice_reads_substring
// origin: languages/csharp/tests/csharp/test_csharp_string_span.rs

using static __Harness;

string s="hello world";
System.ReadOnlySpan<char> span=s.AsSpan(6,5);
__P((span.ToString()).ToString());
__Check("world");

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
