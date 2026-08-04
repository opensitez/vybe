// vybe-test: csharp/csharp_array_2d/two_d_array_row_and_column_lengths
// origin: languages/csharp/tests/csharp/test_csharp_array_2d.rs

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

int[,] m=new int[4,5];
__P((m.GetLength(0)).ToString()); __P((m.GetLength(1)).ToString());
__Check("4\n5");
