// vybe-test: csharp/csharp_index_range_slice/string_range_from_end_indices
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

using static __Harness;

string text="program";
__P((text[^4..^1]).ToString());
__Check("gra");

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
