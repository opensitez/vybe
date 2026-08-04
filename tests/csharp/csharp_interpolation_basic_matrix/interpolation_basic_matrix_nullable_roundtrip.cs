// vybe-test: csharp/csharp_interpolation_basic_matrix/interpolation_basic_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_interpolation_basic_matrix.rs

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

// interpolation_basic_matrix
int? maybe = 112; __P((maybe.HasValue && maybe.Value == 112).ToString());
__Check("True");
