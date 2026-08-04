// vybe-test: csharp/csharp_delegate_invocation_matrix/delegate_invocation_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_delegate_invocation_matrix.rs

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

// delegate_invocation_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[74] = 75; __P((map.ContainsKey(74) && map[74] == 75).ToString());
__Check("True");
