// vybe-test: csharp/csharp_value_task/value_task_from_result_returns_value
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task Run() {
    var vt = System.Threading.Tasks.ValueTask.FromResult(17);
    __Check((await vt).ToString(), "17");
}
Run().Wait();
