// vybe-test: csharp/csharp_rectangular_array_traversal/nested_loops_sum_all_cells_in_rectangular_array
// origin: languages/csharp/tests/csharp/test_csharp_rectangular_array_traversal.rs

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

int[,] grid = {
    { 1, 2, 3 },
    { 4, 5, 6 }
};
int sum = 0;
for (int row = 0; row < grid.GetLength(0); row++) {
    for (int col = 0; col < grid.GetLength(1); col++) {
        sum += grid[row, col];
    }
}
__P((sum).ToString());
__Check("21");
