// vybe-test: csharp/csharp_class_indexers/multi_dimensional_style_dual_indexer_reads_matrix_cell
// origin: languages/csharp/tests/csharp/test_csharp_class_indexers.rs

using static __Harness;

__P((new Matrix()[1, 0]).ToString());
__Check("3");

class Matrix {
    int[,] grid = { { 1, 2 }, { 3, 4 } };
    public int this[int row, int col] { get { return grid[row, col]; } }
}

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
