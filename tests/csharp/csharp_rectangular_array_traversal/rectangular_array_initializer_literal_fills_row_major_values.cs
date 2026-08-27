// vybe-test: csharp/csharp_rectangular_array_traversal/rectangular_array_initializer_literal_fills_row_major_values
// origin: languages/csharp/tests/csharp/test_csharp_rectangular_array_traversal.rs

using static __Harness;

int[,] grid = {
    { 1, 2 },
    { 3, 4 }
}
;
__P((grid[0, 1]).ToString());
__P((grid[1, 0]).ToString());
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
