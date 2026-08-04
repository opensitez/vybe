// vybe-test: csharp/csharp_constructor_null_guard_matrix/constructor_null_guard_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_constructor_null_guard_matrix.rs

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

// constructor_null_guard_matrix
int? maybe = 126; __P((maybe.HasValue && maybe.Value == 126).ToString());
__Check("True");
