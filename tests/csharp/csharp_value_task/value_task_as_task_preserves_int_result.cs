// vybe-test: csharp/csharp_value_task/value_task_as_task_preserves_int_result
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<int> Get() { return 88; }
async System.Threading.Tasks.Task Run() {
    var task = Get().AsTask();
    __Check((await task).ToString(), "88");
}
Run().Wait();
