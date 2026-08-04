// vybe-test: csharp/csharp_task_timeout_matrix/task_timeout_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_task_timeout_matrix.rs

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

// task_timeout_matrix
int? maybe = null; int fallback = maybe ?? 86; __P((fallback == 86).ToString());
__Check("True");
