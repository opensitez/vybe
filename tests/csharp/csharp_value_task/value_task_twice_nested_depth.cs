// vybe-test: csharp/csharp_value_task/value_task_twice_nested_depth
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<int> Deep() { return 1; }
async System.Threading.Tasks.ValueTask<int> Mid() { return await Deep() + 1; }
async System.Threading.Tasks.ValueTask<int> Top() { return await Mid() + 1; }
async System.Threading.Tasks.Task Run() {
    __Check((await Top()).ToString(), "3");
}
Run().Wait();
