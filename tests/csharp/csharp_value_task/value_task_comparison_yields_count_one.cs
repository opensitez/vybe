// vybe-test: csharp/csharp_value_task/value_task_comparison_yields_count_one
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<int> A() { return 10; }
async System.Threading.Tasks.ValueTask<int> B() { return 5; }
async System.Threading.Tasks.Task Run() {
    int count = (await A() > await B()) ? 1 : 0;
    __Check((count).ToString(), "1");
}
Run().Wait();
