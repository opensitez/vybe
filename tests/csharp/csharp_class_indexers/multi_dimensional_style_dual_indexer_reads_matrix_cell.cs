// vybe-test: csharp/csharp_class_indexers/multi_dimensional_style_dual_indexer_reads_matrix_cell
// origin: languages/csharp/tests/csharp/test_csharp_class_indexers.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

class Matrix {
    int[,] grid = { { 1, 2 }, { 3, 4 } };
    public int this[int row, int col] { get { return grid[row, col]; } }
}
__P((new Matrix()[1, 0]).ToString());
__Check("3");
