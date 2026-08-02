// vybe-test: csharp/csharp_async_task/task_from_result_completes_synchronously_with_value
// origin: languages/csharp/tests/csharp/test_csharp_async_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task Run() {
    var t = System.Threading.Tasks.Task.FromResult(42);
    __Check((await t).ToString(), "42");
}
Run().Wait();
