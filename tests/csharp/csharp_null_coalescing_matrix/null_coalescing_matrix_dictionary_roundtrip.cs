// vybe-test: csharp/csharp_null_coalescing_matrix/null_coalescing_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_null_coalescing_matrix.rs

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

// null_coalescing_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[56] = 57; __P((map.ContainsKey(56) && map[56] == 57).ToString());
__Check("True");
