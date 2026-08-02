// vybe-test: csharp/csharp_value_task/value_task_subtraction_chain_count
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<int> Start() { return 50; }
async System.Threading.Tasks.ValueTask<int> Take() { return 8; }
async System.Threading.Tasks.Task Run() {
    __Check((await Start() - await Take()).ToString(), "42");
}
Run().Wait();
