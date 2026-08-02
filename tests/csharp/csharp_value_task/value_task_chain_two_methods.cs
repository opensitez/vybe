// vybe-test: csharp/csharp_value_task/value_task_chain_two_methods
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<int> A() { return 2; }
async System.Threading.Tasks.ValueTask<int> B(int x) { return x + 3; }
async System.Threading.Tasks.Task Run() {
    __Check((await B(await A())).ToString(), "5");
}
Run().Wait();
