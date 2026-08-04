// vybe-test: csharp/csharp_collection_initializer_syntax/jagged_array_initializer_creates_rows_with_varying_lengths
// origin: languages/csharp/tests/csharp/test_csharp_collection_initializer_syntax.rs

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

int[][] grid = {
    new[] { 1, 2 },
    new[] { 3, 4, 5 }
};
__P((grid[1].Length).ToString());
__P((grid[1][2]).ToString());
__Check("3\n5");
