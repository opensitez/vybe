// vybe-test: csharp/csharp_value_task/value_task_from_task_from_result
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

async System.Threading.Tasks.ValueTask<int> ViaTask() {
    return await System.Threading.Tasks.Task.FromResult(21);
}
async System.Threading.Tasks.Task Run() {
    __P((await ViaTask()).ToString());
}
Run().Wait();
__Check("21");
