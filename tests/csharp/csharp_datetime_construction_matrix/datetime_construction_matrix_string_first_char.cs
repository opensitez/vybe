// vybe-test: csharp/csharp_datetime_construction_matrix/datetime_construction_matrix_string_first_char
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
string feature = "datetime_construction_matrix"; __P((feature[0] == feature[0]).ToString());
__Check("True");
