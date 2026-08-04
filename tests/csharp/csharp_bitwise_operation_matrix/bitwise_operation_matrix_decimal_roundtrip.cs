// vybe-test: csharp/csharp_bitwise_operation_matrix/bitwise_operation_matrix_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_bitwise_operation_matrix.rs

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

// bitwise_operation_matrix
decimal amount = 10m; __P(((amount / 2m) * 2m == 10m).ToString());
__Check("True");
