// vybe-test: csharp/csharp_task_combinators/when_any_task_run_vs_from_result
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task Run() {
    var winner = await System.Threading.Tasks.Task.WhenAny(
        System.Threading.Tasks.Task.Run(() => 15),
        System.Threading.Tasks.Task.FromResult(16)
    );
    __Check((winner.Result).ToString(), "15");
}
Run().Wait();
