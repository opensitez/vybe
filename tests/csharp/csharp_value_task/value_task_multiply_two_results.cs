// vybe-test: csharp/csharp_value_task/value_task_multiply_two_results
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<int> Left() { return 6; }
async System.Threading.Tasks.ValueTask<int> Right() { return 7; }
async System.Threading.Tasks.Task Run() {
    __Check((await Left() * await Right()).ToString(), "42");
}
Run().Wait();
