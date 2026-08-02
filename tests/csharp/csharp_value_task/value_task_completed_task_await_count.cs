// vybe-test: csharp/csharp_value_task/value_task_completed_task_await_count
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task Run() {
    await System.Threading.Tasks.ValueTask.CompletedTask;
    __Check((1).ToString(), "1");
}
Run().Wait();
