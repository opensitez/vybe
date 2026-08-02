// vybe-test: csharp/csharp_value_task/sequential_value_task_await_sum_count
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<int> N(int x) { return x; }
async System.Threading.Tasks.Task Run() {
    int total = await N(1) + await N(2) + await N(3);
    __Check((total).ToString(), "6");
}
Run().Wait();
