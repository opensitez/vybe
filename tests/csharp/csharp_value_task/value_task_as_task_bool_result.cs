// vybe-test: csharp/csharp_value_task/value_task_as_task_bool_result
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

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

async System.Threading.Tasks.ValueTask<bool> Yes() { return true; }
async System.Threading.Tasks.Task Run() {
    bool v = await Yes().AsTask();
    __P((v ? 1 : 0).ToString());
}
Run().Wait();
__Check("1");
