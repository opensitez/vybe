// vybe-test: csharp/csharp_async_task_matrix/task_completed_task_reports_completion
// origin: languages/csharp/tests/csharp/test_csharp_async_task_matrix.rs

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

var completed = System.Threading.Tasks.Task.CompletedTask;
__P((completed.IsCompleted).ToString());
__P((completed.IsFaulted).ToString());
__Check("True\nFalse");
