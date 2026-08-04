// vybe-test: csharp/csharp_async_stream_matrix/async_stream_matrix_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_async_stream_matrix.rs

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

// async_stream_matrix
int seed = 111; bool cond = seed % 2 == 0; __P((cond || !cond).ToString());
__Check("True");
