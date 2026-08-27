// vybe-test: csharp/csharp_index_from_end/index_from_end_one_reads_last_element_of_array
// origin: languages/csharp/tests/csharp/test_csharp_index_from_end.rs

using static __Harness;

int[] data = { 10, 20, 30 }
;
__P((data[^1]).ToString());
__Check("30");

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
