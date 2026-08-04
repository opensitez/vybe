// vybe-test: csharp/csharp_jagged_arrays/jagged_array_foreach_over_rows_sums_correctly
// origin: languages/csharp/tests/csharp/test_csharp_jagged_arrays.rs

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

int[][] jag = new[]{ new[]{1,2}, new[]{3,4,5} };
int total=0;
foreach(var row in jag) foreach(var v in row) total+=v;
__P((total).ToString());
__Check("15");
