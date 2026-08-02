// vybe-test: csharp/csharp_task_combinators/when_any_with_task_run_winner_count
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task Run() {
    var winner = await System.Threading.Tasks.Task.WhenAny(
        System.Threading.Tasks.Task.Run(() => 6),
        System.Threading.Tasks.Task.Run(() => 7)
    );
    __Check((winner.Result).ToString(), "6");
}
Run().Wait();
