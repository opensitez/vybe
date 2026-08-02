// vybe-test: csharp/csharp_value_task/async_value_task_with_yield_then_return
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<int> Compute() {
    await System.Threading.Tasks.Task.Yield();
    return 9;
}
async System.Threading.Tasks.Task Run() {
    __Check((await Compute()).ToString(), "9");
}
Run().Wait();
