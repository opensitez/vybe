// vybe-test: csharp/csharp_task_combinators/when_all_single_task_returns_singleton
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task<int> Solo() { return 11; }
async System.Threading.Tasks.Task Run() {
    var results = await System.Threading.Tasks.Task.WhenAll(Solo());
    __Check((results[0]).ToString(), "11");
}
Run().Wait();
