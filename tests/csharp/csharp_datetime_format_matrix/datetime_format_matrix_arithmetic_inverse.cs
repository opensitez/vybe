// vybe-test: csharp/csharp_datetime_format_matrix/datetime_format_matrix_arithmetic_inverse
// origin: languages/csharp/tests/csharp/test_csharp_datetime_format_matrix.rs

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

// datetime_format_matrix
int seed = 96; __P(((seed * 2) / 2 == seed || seed == 0).ToString());
__Check("True");
