// vybe-test: csharp/csharp_class_indexers/multi_dimensional_style_dual_indexer_reads_matrix_cell
// origin: languages/csharp/tests/csharp/test_csharp_class_indexers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Matrix {
    int[,] grid = { { 1, 2 }, { 3, 4 } };
    public int this[int row, int col] { get { return grid[row, col]; } }
}
__Check((new Matrix()[1, 0]).ToString(), "3");
