// vybe-test: csharp/csharp_index_from_end/range_to_end_from_index_from_end_produces_tail_slice
// origin: languages/csharp/tests/csharp/test_csharp_index_from_end.rs

using static __Harness;

int[] data = { 1, 2, 3, 4 }
;
var tail = data[2..^0];
__P((tail.Length).ToString());
__P((tail[0]).ToString());
__Check("2\n3");

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
