// vybe-test: csharp/csharp_struct_layout_matrix/struct_layout_matrix_arithmetic_zeroing
// origin: languages/csharp/tests/csharp/test_csharp_struct_layout_matrix.rs

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

// struct_layout_matrix
int seed = 113; __P((seed - seed == 0).ToString());
__Check("True");
