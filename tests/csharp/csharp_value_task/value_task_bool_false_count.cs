// vybe-test: csharp/csharp_value_task/value_task_bool_false_count
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<bool> No() { return false; }
async System.Threading.Tasks.Task Run() {
    bool v = await No();
    __Check((v ? 1 : 0).ToString(), "0");
}
Run().Wait();
