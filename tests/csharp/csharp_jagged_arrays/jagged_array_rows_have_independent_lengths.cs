// vybe-test: csharp/csharp_jagged_arrays/jagged_array_rows_have_independent_lengths
// origin: languages/csharp/tests/csharp/test_csharp_jagged_arrays.rs

using static __Harness;

int[][] jag = new int[3][];
jag[0] = new int[]{1}
;
jag[1] = new int[]{2,3}
;
jag[2] = new int[]{4,5,6}
;
__P((jag[2].Length).ToString());
__Check("3");

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
