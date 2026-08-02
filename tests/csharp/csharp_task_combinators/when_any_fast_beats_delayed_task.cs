// vybe-test: csharp/csharp_task_combinators/when_any_fast_beats_delayed_task
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task<int> Fast() { return 1; }
async System.Threading.Tasks.Task<int> Slow() {
    await System.Threading.Tasks.Task.Delay(1000);
    return 2;
}
async System.Threading.Tasks.Task Run() {
    var winner = await System.Threading.Tasks.Task.WhenAny(Fast(), Slow());
    __Check((winner.Result).ToString(), "1");
}
Run().Wait();
