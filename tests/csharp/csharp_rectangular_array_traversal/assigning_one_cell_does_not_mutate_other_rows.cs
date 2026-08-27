// vybe-test: csharp/csharp_rectangular_array_traversal/assigning_one_cell_does_not_mutate_other_rows
// origin: languages/csharp/tests/csharp/test_csharp_rectangular_array_traversal.rs

using static __Harness;

int[,] grid = {
    { 10, 20 },
    { 30, 40 }
}
;
grid[1, 1] = 99;
__P((grid[0, 1]).ToString());
__P((grid[1, 1]).ToString());
__Check("20\n99");

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
