// vybe-test: csharp/csharp_task_combinators/when_any_picks_from_result_fast_task
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task Run() {
    var winner = await System.Threading.Tasks.Task.WhenAny(
        System.Threading.Tasks.Task.FromResult(3),
        System.Threading.Tasks.Task.FromResult(9)
    );
    __Check((winner.Result).ToString(), "3");
}
Run().Wait();
