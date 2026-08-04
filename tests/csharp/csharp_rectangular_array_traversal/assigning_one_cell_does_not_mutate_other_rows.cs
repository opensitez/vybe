// vybe-test: csharp/csharp_rectangular_array_traversal/assigning_one_cell_does_not_mutate_other_rows
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
    { 10, 20 },
    { 30, 40 }
};
grid[1, 1] = 99;
__P((grid[0, 1]).ToString());
__P((grid[1, 1]).ToString());
__Check("20\n99");
