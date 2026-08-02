// vybe-test: csharp/csharp_value_task/value_task_expression_bodied_async
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<int> Get() => 55;
async System.Threading.Tasks.Task Run() {
    __Check((await Get()).ToString(), "55");
}
Run().Wait();
