// vybe-test: csharp/csharp_span_indexing/span_from_array_slice_reads_correct_elements
// origin: languages/csharp/tests/csharp/test_csharp_span_indexing.rs

using static __Harness;

int[] data = { 10, 20, 30, 40, 50 }
;
var span = new System.Span<int>(data, 1, 3);
__P((span[0]).ToString());
__P((span[2]).ToString());
__Check("20\n40");

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
