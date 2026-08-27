// vybe-test: csharp/csharp_stackalloc_span/stackalloc_int_zero_length_span_has_zero_length
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

using static __Harness;

Span<int> span = stackalloc int[4];
span[0] = 10;
span[1] = 20;
__P(span[0].ToString());
__P(span[1].ToString());
__Check("10\n20");
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
