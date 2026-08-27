// vybe-test: csharp/csharp_jagged_arrays/jagged_array_foreach_over_rows_sums_correctly
// origin: languages/csharp/tests/csharp/test_csharp_jagged_arrays.rs

using static __Harness;

int[][] jag = new[]{ new[]{1,2}, new[]{3,4,5} }
;
int total=0;
foreach(var row in jag) foreach(var v in row) total+=v;
__P((total).ToString());
__Check("15");

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
