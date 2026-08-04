// vybe-test: csharp/csharp_jagged_arrays/two_dimensional_array_element_set_and_read
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

int[,] m = new int[2,2];
m[0,1] = 7;
__P((m[0,1]).ToString());
__Check("7");
