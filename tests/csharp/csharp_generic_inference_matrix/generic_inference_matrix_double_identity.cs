// vybe-test: csharp/csharp_generic_inference_matrix/generic_inference_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_matrix.rs

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

// generic_inference_matrix
double seed = 81; __P(((seed + 0.5 - 0.5) == seed).ToString());
__Check("True");
