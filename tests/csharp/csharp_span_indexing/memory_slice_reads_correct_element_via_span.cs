// vybe-test: csharp/csharp_span_indexing/memory_slice_reads_correct_element_via_span
// origin: languages/csharp/tests/csharp/test_csharp_span_indexing.rs

using static __Harness;

var memory = new System.Memory<int>(new int[] { 5, 6, 7 });
__P((memory.Span[2]).ToString());
__Check("7");

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
