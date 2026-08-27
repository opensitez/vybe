// vybe-test: csharp/csharp_jagged_arrays/array_rank_is_one_for_flat_array
// origin: languages/csharp/tests/csharp/test_csharp_jagged_arrays.rs

using static __Harness;

int[] a = new int[5];
__P((a.Rank).ToString());
__Check("1");

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
