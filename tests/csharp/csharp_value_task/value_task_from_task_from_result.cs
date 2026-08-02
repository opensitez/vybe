// vybe-test: csharp/csharp_value_task/value_task_from_task_from_result
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<int> ViaTask() {
    return await System.Threading.Tasks.Task.FromResult(21);
}
async System.Threading.Tasks.Task Run() {
    __Check((await ViaTask()).ToString(), "21");
}
Run().Wait();
