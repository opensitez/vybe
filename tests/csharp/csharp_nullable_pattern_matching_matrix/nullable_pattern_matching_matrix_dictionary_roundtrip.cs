// vybe-test: csharp/csharp_nullable_pattern_matching_matrix/nullable_pattern_matching_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_nullable_pattern_matching_matrix.rs

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

// nullable_pattern_matching_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[125] = 126; __P((map.ContainsKey(125) && map[125] == 126).ToString());
__Check("True");
