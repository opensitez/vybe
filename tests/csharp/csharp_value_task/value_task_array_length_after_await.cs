// vybe-test: csharp/csharp_value_task/value_task_array_length_after_await
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

async System.Threading.Tasks.ValueTask<int[]> Get() {
    return new int[] { 1, 2, 3, 4, 5 };
}
async System.Threading.Tasks.Task Run() {
    int[] arr = await Get();
    __P((arr.Length).ToString());
}
Run().Wait();
__Check("5");
