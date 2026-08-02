// vybe-test: csharp/csharp_value_task/nested_value_task_await_chain
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<int> Inner() { return 5; }
async System.Threading.Tasks.ValueTask<int> Outer() { return await Inner() + 1; }
async System.Threading.Tasks.Task Run() {
    __Check((await Outer()).ToString(), "6");
}
Run().Wait();
