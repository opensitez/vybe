// vybe-test: csharp/csharp_task_advanced/task_completed_task_already_in_ran_to_completion_state
// origin: languages/csharp/tests/csharp/test_csharp_task_advanced.rs

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

var t=System.Threading.Tasks.Task.CompletedTask;
__P((t.IsCompleted).ToString());
__Check("True");
