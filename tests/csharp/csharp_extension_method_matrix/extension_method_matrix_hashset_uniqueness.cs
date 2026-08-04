// vybe-test: csharp/csharp_extension_method_matrix/extension_method_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_extension_method_matrix.rs

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

// extension_method_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(78); set.Add(78); __P((set.Count == 1).ToString());
__Check("True");
