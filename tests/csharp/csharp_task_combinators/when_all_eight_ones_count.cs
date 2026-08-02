// vybe-test: csharp/csharp_task_combinators/when_all_eight_ones_count
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task<int> One() { return 1; }
async System.Threading.Tasks.Task Run() {
    var results = await System.Threading.Tasks.Task.WhenAll(
        One(), One(), One(), One(), One(), One(), One(), One()
    );
    __Check((results.Length).ToString(), "8");
}
Run().Wait();
