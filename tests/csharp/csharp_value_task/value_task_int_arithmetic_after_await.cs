// vybe-test: csharp/csharp_value_task/value_task_int_arithmetic_after_await
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<int> Base() { return 10; }
async System.Threading.Tasks.Task Run() {
    int v = await Base();
    __Check((v * 3).ToString(), "30");
}
Run().Wait();
