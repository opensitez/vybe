// vybe-test: csharp/csharp_async_enumerator_matrix/async_enumerator_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_async_enumerator_matrix.rs

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

// async_enumerator_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[116] = 117; __P((map.ContainsKey(116) && map[116] == 117).ToString());
__Check("True");
