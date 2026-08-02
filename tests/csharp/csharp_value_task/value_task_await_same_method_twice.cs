// vybe-test: csharp/csharp_value_task/value_task_await_same_method_twice
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<int> Constant() { return 6; }
async System.Threading.Tasks.Task Run() {
    int a = await Constant();
    int b = await Constant();
    __Check((a + b).ToString(), "12");
}
Run().Wait();
