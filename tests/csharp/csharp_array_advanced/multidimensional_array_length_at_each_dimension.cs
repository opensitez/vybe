// vybe-test: csharp/csharp_array_advanced/multidimensional_array_length_at_each_dimension
// origin: languages/csharp/tests/csharp/test_csharp_array_advanced.rs

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

int[,] grid=new int[3,4];
__P((grid.GetLength(0)).ToString()); __P((grid.GetLength(1)).ToString());
__Check("3\n4");
