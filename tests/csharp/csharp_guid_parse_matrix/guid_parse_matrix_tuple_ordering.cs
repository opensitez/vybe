// vybe-test: csharp/csharp_guid_parse_matrix/guid_parse_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_guid_parse_matrix.rs

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

// guid_parse_matrix
var tuple = (left: 97, right: 98); __P((tuple.left < tuple.right).ToString());
__Check("True");
