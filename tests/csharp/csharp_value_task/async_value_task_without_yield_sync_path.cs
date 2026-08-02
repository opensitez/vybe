// vybe-test: csharp/csharp_value_task/async_value_task_without_yield_sync_path
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<int> Compute() { return 3 + 4; }
async System.Threading.Tasks.Task Run() {
    __Check((await Compute()).ToString(), "7");
}
Run().Wait();
