// vybe-test: csharp/csharp_value_task/generic_value_task_method_count
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<T> Identity<T>(T value) { return value; }
async System.Threading.Tasks.Task Run() {
    int count = await Identity(4);
    __Check((count).ToString(), "4");
}
Run().Wait();
