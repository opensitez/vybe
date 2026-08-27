// vybe-test: csharp/csharp_deconstruction/nested_tuple_deconstruction_reads_inner_values
// origin: languages/csharp/tests/csharp/test_csharp_deconstruction.rs

using static __Harness;

var ((x, y), label) = ((5, 6), "pt");
__P((label).ToString());
__P((x + y).ToString());
__Check("pt\n11");

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
