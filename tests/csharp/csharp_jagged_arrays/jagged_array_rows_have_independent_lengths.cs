// vybe-test: csharp/csharp_jagged_arrays/jagged_array_rows_have_independent_lengths
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

int[][] jag = new int[3][];
jag[0] = new int[]{1};
jag[1] = new int[]{2,3};
jag[2] = new int[]{4,5,6};
__P((jag[2].Length).ToString());
__Check("3");
