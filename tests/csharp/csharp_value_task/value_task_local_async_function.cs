// vybe-test: csharp/csharp_value_task/value_task_local_async_function
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task Run() {
    async System.Threading.Tasks.ValueTask<int> Local() { return 12; }
    __Check((await Local()).ToString(), "12");
}
Run().Wait();
