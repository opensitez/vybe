// vybe-test: csharp/csharp_value_task/value_task_array_length_after_await
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<int[]> Get() {
    return new int[] { 1, 2, 3, 4, 5 };
}
async System.Threading.Tasks.Task Run() {
    int[] arr = await Get();
    __Check((arr.Length).ToString(), "5");
}
Run().Wait();
