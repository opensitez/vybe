// vybe-test: csharp/csharp_path_api_matrix/path_api_matrix_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_path_api_matrix.rs

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

// path_api_matrix
string feature = "path_api_matrix"; __P((feature.Substring(0, 1) == feature[0].ToString()).ToString());
__Check("True");
