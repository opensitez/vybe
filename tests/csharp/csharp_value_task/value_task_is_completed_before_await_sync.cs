// vybe-test: csharp/csharp_value_task/value_task_is_completed_before_await_sync
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

async System.Threading.Tasks.ValueTask<int> Sync() { return 11; }
async System.Threading.Tasks.Task Run() {
    var vt = Sync();
    int count = vt.IsCompleted ? 1 : 0;
    __P((count).ToString());
}
Run().Wait();
__Check("1");
