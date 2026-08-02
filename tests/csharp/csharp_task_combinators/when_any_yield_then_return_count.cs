// vybe-test: csharp/csharp_task_combinators/when_any_yield_then_return_count
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task<int> Yielded() {
    await System.Threading.Tasks.Task.Yield();
    return 8;
}
async System.Threading.Tasks.Task Run() {
    var winner = await System.Threading.Tasks.Task.WhenAny(Yielded(), Yielded());
    __Check((winner.Result).ToString(), "8");
}
Run().Wait();
