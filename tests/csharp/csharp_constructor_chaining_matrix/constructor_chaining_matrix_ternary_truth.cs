// vybe-test: csharp/csharp_constructor_chaining_matrix/constructor_chaining_matrix_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chaining_matrix.rs

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

// constructor_chaining_matrix
int seed = 68; bool cond = seed % 2 == 0; __P((cond || !cond).ToString());
__Check("True");
