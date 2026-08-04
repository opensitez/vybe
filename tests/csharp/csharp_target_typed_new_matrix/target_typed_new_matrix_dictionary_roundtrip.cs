// vybe-test: csharp/csharp_target_typed_new_matrix/target_typed_new_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_matrix.rs

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

// target_typed_new_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[107] = 108; __P((map.ContainsKey(107) && map[107] == 108).ToString());
__Check("True");
