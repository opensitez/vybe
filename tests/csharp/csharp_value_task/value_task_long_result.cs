// vybe-test: csharp/csharp_value_task/value_task_long_result
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<long> Get() { return 1000000L; }
async System.Threading.Tasks.Task Run() {
    __Check((await Get()).ToString(), "1000000");
}
Run().Wait();
