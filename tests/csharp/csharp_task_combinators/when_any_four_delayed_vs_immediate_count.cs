// vybe-test: csharp/csharp_task_combinators/when_any_four_delayed_vs_immediate_count
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task<int> Now() { return 2; }
async System.Threading.Tasks.Task<int> Later() {
    await System.Threading.Tasks.Task.Delay(200);
    return 3;
}
async System.Threading.Tasks.Task Run() {
    var winner = await System.Threading.Tasks.Task.WhenAny(Now(), Later(), Later(), Later());
    __Check((winner.Result).ToString(), "2");
}
Run().Wait();
