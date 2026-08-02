// vybe-test: csharp/csharp_value_task/value_task_configure_await_false
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<int> Compute() {
    await System.Threading.Tasks.Task.Yield().ConfigureAwait(false);
    return 33;
}
async System.Threading.Tasks.Task Run() {
    __Check((await Compute()).ToString(), "33");
}
Run().Wait();
