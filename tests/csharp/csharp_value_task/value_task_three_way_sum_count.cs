// vybe-test: csharp/csharp_value_task/value_task_three_way_sum_count
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<int> N(int v) { return v; }
async System.Threading.Tasks.Task Run() {
    int count = await N(4) + await N(5) + await N(6);
    __Check((count).ToString(), "15");
}
Run().Wait();
