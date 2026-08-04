// vybe-test: csharp/csharp_math_trigonometry_matrix/math_trigonometry_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_math_trigonometry_matrix.rs

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

// math_trigonometry_matrix
var values = new System.Collections.Generic.List<int> { 102, 103, 102 }; __P((values.Count == 3).ToString());
__Check("True");
