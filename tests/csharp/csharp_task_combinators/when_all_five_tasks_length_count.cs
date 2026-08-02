// vybe-test: csharp/csharp_task_combinators/when_all_five_tasks_length_count
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task<int> N(int v) { return v; }
async System.Threading.Tasks.Task Run() {
    var results = await System.Threading.Tasks.Task.WhenAll(N(1), N(2), N(3), N(4), N(5));
    __Check((results.Length).ToString(), "5");
}
Run().Wait();
