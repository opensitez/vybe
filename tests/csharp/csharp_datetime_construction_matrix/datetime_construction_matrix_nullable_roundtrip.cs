// vybe-test: csharp/csharp_datetime_construction_matrix/datetime_construction_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_datetime_construction_matrix.rs

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

// datetime_construction_matrix
int? maybe = 94; __P((maybe.HasValue && maybe.Value == 94).ToString());
__Check("True");
