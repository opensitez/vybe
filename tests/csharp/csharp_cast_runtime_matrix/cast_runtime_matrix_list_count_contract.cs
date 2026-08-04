// vybe-test: csharp/csharp_cast_runtime_matrix/cast_runtime_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_cast_runtime_matrix.rs

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

// cast_runtime_matrix
var values = new System.Collections.Generic.List<int> { 61, 62, 61 }; __P((values.Count == 3).ToString());
__Check("True");
