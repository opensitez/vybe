// vybe-test: csharp/csharp_value_task/value_task_delay_via_as_task_count
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<int> Delayed() {
    await System.Threading.Tasks.Task.Delay(0).ConfigureAwait(false);
    return 2;
}
async System.Threading.Tasks.Task Run() {
    var task = Delayed().AsTask();
    __Check((await task).ToString(), "2");
}
Run().Wait();
