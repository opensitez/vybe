// vybe-test: csharp/csharp_value_task/value_task_delay_via_as_task_count
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

async System.Threading.Tasks.ValueTask<int> Delayed() {
    await System.Threading.Tasks.Task.Delay(0).ConfigureAwait(false);
    return 2;
}
async System.Threading.Tasks.Task Run() {
    var task = Delayed().AsTask();
    __P((await task).ToString());
}
Run().Wait();
__Check("2");
