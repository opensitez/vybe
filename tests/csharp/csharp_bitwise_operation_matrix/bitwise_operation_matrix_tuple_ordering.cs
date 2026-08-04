// vybe-test: csharp/csharp_bitwise_operation_matrix/bitwise_operation_matrix_tuple_ordering
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
var tuple = (left: 104, right: 105); __P((tuple.left < tuple.right).ToString());
__Check("True");
